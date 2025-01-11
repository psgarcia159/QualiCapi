VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form TelaDuplicatas 
   Caption         =   "Manutenção e Consulta de Duplicatas"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   1470
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   12945
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel PanBaixaParcial 
      Height          =   3165
      Left            =   2520
      TabIndex        =   69
      Top             =   600
      Visible         =   0   'False
      Width           =   7875
      _Version        =   65536
      _ExtentX        =   13891
      _ExtentY        =   5583
      _StockProps     =   15
      Caption         =   " Baixa Parcial"
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
      Enabled         =   0   'False
      Begin Threed.SSPanel PanResiduo 
         Height          =   1125
         Left            =   105
         TabIndex        =   84
         Top             =   1545
         Width           =   7665
         _Version        =   65536
         _ExtentX        =   13520
         _ExtentY        =   1984
         _StockProps     =   15
         Caption         =   " Resíduo"
         ForeColor       =   4210688
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
         BevelOuter      =   0
         BevelInner      =   2
         Alignment       =   0
         Begin VB.ComboBox CboTipoPagRecResiduo 
            Height          =   315
            ItemData        =   "TelaDuplicatas.frx":0000
            Left            =   4410
            List            =   "TelaDuplicatas.frx":0019
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   330
            Width           =   3150
         End
         Begin VB.TextBox TxtCodCCorrenteResiduo 
            Height          =   330
            Left            =   1485
            MaxLength       =   4
            TabIndex        =   80
            Top             =   675
            Width           =   450
         End
         Begin MSDataListLib.DataCombo CboCCorrenteResiduo 
            Bindings        =   "TelaDuplicatas.frx":0073
            Height          =   330
            Left            =   1980
            TabIndex        =   81
            Top             =   675
            Width           =   5580
            _ExtentX        =   9843
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
         Begin MSComCtl2.DTPicker DtpVencResiduo 
            Height          =   315
            Left            =   1485
            TabIndex        =   78
            Top             =   330
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MM/yy"
            Format          =   81657857
            CurrentDate     =   37460
         End
         Begin VB.Label LblTipoResiduo 
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
            Left            =   2790
            TabIndex        =   87
            Top             =   375
            Width           =   1590
         End
         Begin VB.Label LblCCorrenteResiduo 
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
            Left            =   105
            TabIndex        =   86
            Top             =   720
            Width           =   1365
         End
         Begin VB.Label LblVencResiduo 
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
            Left            =   330
            TabIndex        =   85
            Top             =   360
            Width           =   1125
         End
      End
      Begin Threed.SSPanel PanParcial 
         Height          =   1170
         Left            =   150
         TabIndex        =   88
         Top             =   360
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   2064
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
         Begin VB.TextBox TxtCodTipoDoc 
            Height          =   315
            Left            =   1050
            MaxLength       =   4
            TabIndex        =   70
            Top             =   75
            Width           =   495
         End
         Begin VB.TextBox TxtCodCCorrenteParcial 
            Height          =   330
            Left            =   1500
            MaxLength       =   4
            TabIndex        =   76
            Top             =   765
            Width           =   450
         End
         Begin VB.ComboBox CboTipoPagRecParcial 
            Height          =   315
            ItemData        =   "TelaDuplicatas.frx":0092
            Left            =   2940
            List            =   "TelaDuplicatas.frx":00AB
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   420
            Width           =   2430
         End
         Begin VB.TextBox TxtNumDocParcial 
            Height          =   315
            Left            =   6390
            MaxLength       =   7
            TabIndex        =   75
            Top             =   420
            Width           =   1110
         End
         Begin VB.TextBox TxtValorParcial 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5775
            TabIndex        =   72
            Top             =   75
            Width           =   1725
         End
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   330
            Left            =   2475
            Top             =   795
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
         Begin MSDataListLib.DataCombo CboCCorrenteParcial 
            Bindings        =   "TelaDuplicatas.frx":0105
            Height          =   330
            Left            =   1980
            TabIndex        =   77
            Top             =   765
            Width           =   5535
            _ExtentX        =   9763
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
         Begin MSComCtl2.DTPicker DtpPagtoParcial 
            Height          =   315
            Left            =   1155
            TabIndex        =   73
            Top             =   420
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MM/yy"
            Format          =   81657857
            CurrentDate     =   37460
         End
         Begin MSAdodcLib.Adodc DatTipoDoc 
            Height          =   330
            Left            =   2985
            Top             =   75
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
            Bindings        =   "TelaDuplicatas.frx":0124
            Height          =   315
            Left            =   1575
            TabIndex        =   71
            Top             =   75
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "tido_tx_descricao"
            BoundColumn     =   "tido_cd_tipodoc"
            Text            =   "CboTipoDoc"
         End
         Begin VB.Label LblTipoDoc 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tipo Doc.:"
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
            Left            =   105
            TabIndex        =   94
            Top             =   120
            Width           =   930
         End
         Begin VB.Label LblCCorrenteParcial 
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
            Height          =   300
            Left            =   105
            TabIndex        =   93
            Top             =   795
            Width           =   1365
         End
         Begin VB.Label LblDtPagtoParcial 
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
            Left            =   30
            TabIndex        =   92
            Top             =   450
            Width           =   1110
         End
         Begin VB.Label LblNumDocParcial 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nº Doc.:"
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
            Left            =   5490
            TabIndex        =   91
            Top             =   450
            Width           =   855
         End
         Begin VB.Label LblTipoPagRecParcial 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Height          =   255
            Left            =   2445
            TabIndex        =   90
            Top             =   450
            Width           =   465
         End
         Begin VB.Label LblValorParcial 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Left            =   5205
            TabIndex        =   89
            Top             =   120
            Width           =   555
         End
      End
      Begin VB.CommandButton CmdConfimarParcial 
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
         Left            =   5640
         TabIndex        =   82
         Top             =   2685
         Width           =   1005
      End
      Begin VB.CommandButton CmdRetornarParcial 
         Caption         =   "Retornar"
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
         Left            =   6720
         TabIndex        =   83
         Top             =   2685
         Width           =   1005
      End
   End
   Begin Threed.SSPanel PanEstorno 
      Height          =   1410
      Left            =   1320
      TabIndex        =   95
      Top             =   2280
      Visible         =   0   'False
      Width           =   3270
      _Version        =   65536
      _ExtentX        =   5768
      _ExtentY        =   2487
      _StockProps     =   15
      Caption         =   " Estorno"
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
      Enabled         =   0   'False
      Begin VB.CommandButton CmdConfirmaEstorno 
         Caption         =   "&Confirmar"
         Height          =   255
         Left            =   1395
         TabIndex        =   97
         Top             =   1020
         Width           =   825
      End
      Begin VB.CommandButton CmdRetornaEstorno 
         Caption         =   "Retornar"
         Height          =   255
         Left            =   2265
         TabIndex        =   96
         Top             =   1020
         Width           =   870
      End
      Begin MSComCtl2.DTPicker DtpDtEstorno 
         Height          =   315
         Left            =   1455
         TabIndex        =   98
         Top             =   510
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         DateIsNull      =   -1  'True
         Format          =   81657857
         CurrentDate     =   37460
      End
      Begin VB.Label LblEstorno 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Data Estorno:"
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
         TabIndex        =   99
         Top             =   540
         Width           =   1230
      End
   End
   Begin Threed.SSPanel PanCopia 
      Height          =   2190
      Left            =   1200
      TabIndex        =   43
      Top             =   1440
      Visible         =   0   'False
      Width           =   7170
      _Version        =   65536
      _ExtentX        =   12647
      _ExtentY        =   3863
      _StockProps     =   15
      Caption         =   "Cópia de Cheque"
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
      Begin VB.TextBox TxtCCorrente2 
         Height          =   330
         Left            =   1545
         MaxLength       =   4
         TabIndex        =   47
         Top             =   1290
         Width           =   450
      End
      Begin VB.CommandButton CmdRetornar2 
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
         Left            =   5910
         TabIndex        =   50
         Top             =   1710
         Width           =   1110
      End
      Begin VB.CommandButton CmdImpCopia 
         Caption         =   "I&mprimir"
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
         Left            =   4725
         TabIndex        =   49
         Top             =   1710
         Width           =   1110
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nº do Cheque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   1605
         TabIndex        =   44
         Top             =   480
         Width           =   4110
         Begin VB.TextBox TxtNroCheque3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2820
            TabIndex        =   46
            Top             =   270
            Width           =   855
         End
         Begin VB.TextBox TxtNroCheque2 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1125
            TabIndex        =   45
            Top             =   270
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "De:"
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
            Left            =   690
            TabIndex        =   52
            Top             =   285
            Width           =   390
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "Até:"
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
            Left            =   2415
            TabIndex        =   51
            Top             =   285
            Width           =   375
         End
      End
      Begin MSDataListLib.DataCombo CboCCorrente2 
         Bindings        =   "TelaDuplicatas.frx":013D
         Height          =   330
         Left            =   2010
         TabIndex        =   48
         Top             =   1290
         Width           =   5040
         _ExtentX        =   8890
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
      Begin VB.Label Label7 
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
         Height          =   300
         Left            =   150
         TabIndex        =   53
         Top             =   1320
         Width           =   1365
      End
   End
   Begin Threed.SSPanel panCancelar 
      Height          =   1290
      Left            =   6480
      TabIndex        =   101
      Top             =   2400
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
         TabIndex        =   103
         Top             =   960
         Width           =   870
      End
      Begin VB.CommandButton cmdConfirmarCancelamento 
         Caption         =   "&Confirmar"
         Height          =   255
         Left            =   1920
         TabIndex        =   102
         Top             =   960
         Width           =   825
      End
      Begin MSComCtl2.DTPicker dtpCancelamento 
         Height          =   315
         Left            =   2400
         TabIndex        =   104
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         DateIsNull      =   -1  'True
         Format          =   81657857
         CurrentDate     =   37460
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
         TabIndex        =   105
         Top             =   540
         Width           =   2175
      End
   End
   Begin Threed.SSPanel PanDuplicata 
      Height          =   1770
      Left            =   5280
      TabIndex        =   58
      Top             =   1440
      Visible         =   0   'False
      Width           =   7785
      _Version        =   65536
      _ExtentX        =   13732
      _ExtentY        =   3122
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
      BorderWidth     =   2
      BevelInner      =   1
      Alignment       =   0
      Enabled         =   0   'False
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
         Height          =   285
         Left            =   5610
         TabIndex        =   66
         Top             =   1365
         Width           =   1005
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
         Height          =   285
         Left            =   6675
         TabIndex        =   67
         Top             =   1365
         Width           =   1005
      End
      Begin Threed.SSPanel PanPagto2 
         Height          =   1050
         Left            =   60
         TabIndex        =   59
         Top             =   60
         Width           =   7665
         _Version        =   65536
         _ExtentX        =   13520
         _ExtentY        =   1852
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
         BevelInner      =   2
         Begin VB.TextBox TxtCCorrente3 
            Height          =   330
            Left            =   1485
            MaxLength       =   4
            TabIndex        =   64
            Top             =   510
            Width           =   450
         End
         Begin VB.ComboBox CboTipoPagto2 
            Height          =   315
            ItemData        =   "TelaDuplicatas.frx":015C
            Left            =   4410
            List            =   "TelaDuplicatas.frx":0175
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   135
            Width           =   3135
         End
         Begin MSDataListLib.DataCombo CboCCorrente3 
            Bindings        =   "TelaDuplicatas.frx":01CF
            Height          =   330
            Left            =   1980
            TabIndex        =   65
            Top             =   510
            Width           =   5565
            _ExtentX        =   9816
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
         Begin MSComCtl2.DTPicker DtpVencimento 
            Height          =   315
            Left            =   1485
            TabIndex        =   62
            Top             =   135
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   81657857
            CurrentDate     =   37460
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
            Left            =   330
            TabIndex        =   68
            Top             =   165
            Width           =   1125
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
            Left            =   105
            TabIndex        =   61
            Top             =   555
            Width           =   1365
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
            Left            =   2790
            TabIndex        =   60
            Top             =   180
            Width           =   1590
         End
      End
   End
   Begin Threed.SSPanel PanPagamento 
      Height          =   2535
      Left            =   1920
      TabIndex        =   25
      Top             =   1440
      Visible         =   0   'False
      Width           =   8040
      _Version        =   65536
      _ExtentX        =   14182
      _ExtentY        =   4471
      _StockProps     =   15
      Caption         =   "Pagamento"
      ForeColor       =   12582912
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
      Enabled         =   0   'False
      Begin MSAdodcLib.Adodc DatFormaPagamento 
         Height          =   330
         Left            =   2040
         Top             =   1440
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
      Begin VB.TextBox TxtCCorrente 
         Height          =   330
         Left            =   1650
         MaxLength       =   4
         TabIndex        =   28
         Top             =   990
         Width           =   450
      End
      Begin VB.ComboBox CboTipoPag 
         Height          =   315
         ItemData        =   "TelaDuplicatas.frx":01EE
         Left            =   3675
         List            =   "TelaDuplicatas.frx":0207
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   510
         Width           =   2400
      End
      Begin MSAdodcLib.Adodc DatContaCorrente 
         Height          =   330
         Left            =   2520
         Top             =   990
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
      Begin VB.CommandButton CmdRetornar 
         Caption         =   "Retornar"
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
         Left            =   6800
         TabIndex        =   31
         Top             =   1920
         Width           =   1005
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
         Height          =   345
         Left            =   5700
         TabIndex        =   30
         Top             =   1920
         Width           =   1005
      End
      Begin VB.TextBox TxtNumDoc 
         Height          =   315
         Left            =   6915
         MaxLength       =   7
         TabIndex        =   27
         Top             =   510
         Width           =   945
      End
      Begin MSDataListLib.DataCombo CboCCorrente 
         Bindings        =   "TelaDuplicatas.frx":0261
         Height          =   330
         Left            =   2145
         TabIndex        =   29
         Top             =   990
         Width           =   5715
         _ExtentX        =   10081
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
      Begin MSComCtl2.DTPicker DtpDtPagto 
         Height          =   315
         Left            =   1680
         TabIndex        =   24
         Top             =   510
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         DateIsNull      =   -1  'True
         Format          =   81657857
         CurrentDate     =   37460
      End
      Begin MSDataListLib.DataCombo CboFormaPagamento 
         Bindings        =   "TelaDuplicatas.frx":0280
         Height          =   330
         Left            =   1650
         TabIndex        =   127
         Top             =   1440
         Width           =   6200
         _ExtentX        =   10927
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
      Begin VB.Label Label1 
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
         Left            =   150
         TabIndex        =   128
         Top             =   1455
         Width           =   1450
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
         Height          =   300
         Left            =   220
         TabIndex        =   42
         Top             =   1020
         Width           =   1365
      End
      Begin VB.Label Label3 
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
         Left            =   350
         TabIndex        =   35
         Top             =   540
         Width           =   1230
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nº Doc.:"
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
         Left            =   6030
         TabIndex        =   34
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Labe1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   255
         Left            =   3180
         TabIndex        =   33
         Top             =   540
         Width           =   465
      End
   End
   Begin Threed.SSPanel PanCheque 
      Height          =   2580
      Left            =   480
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   7305
      _Version        =   65536
      _ExtentX        =   12885
      _ExtentY        =   4551
      _StockProps     =   15
      Caption         =   "Painel de Impressão do Cheque"
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
      Begin VB.TextBox TxtCodBanco 
         Height          =   315
         Left            =   1485
         MaxLength       =   4
         TabIndex        =   19
         Top             =   1230
         Width           =   525
      End
      Begin MSAdodcLib.Adodc Datnominal 
         Height          =   330
         Left            =   2880
         Top             =   1245
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
         Caption         =   "Datnominal"
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
      Begin VB.CommandButton CmdImpCheque 
         Caption         =   "&Imprimir"
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
         Left            =   4860
         TabIndex        =   22
         Top             =   2085
         Width           =   1110
      End
      Begin VB.CommandButton CmdRetorna 
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
         Left            =   6060
         TabIndex        =   23
         Top             =   2085
         Width           =   1110
      End
      Begin VB.TextBox TxtVerso 
         Height          =   300
         Left            =   1485
         MaxLength       =   50
         TabIndex        =   21
         Top             =   1605
         Width           =   5685
      End
      Begin VB.TextBox TxtNroCheque 
         Height          =   300
         Left            =   1485
         MaxLength       =   9
         TabIndex        =   17
         Top             =   540
         Width           =   900
      End
      Begin VB.TextBox TxtCidade 
         Height          =   300
         Left            =   1485
         MaxLength       =   30
         TabIndex        =   18
         Top             =   880
         Width           =   2550
      End
      Begin MSDataListLib.DataCombo CboNominal 
         Bindings        =   "TelaDuplicatas.frx":02A0
         Height          =   315
         Left            =   2040
         TabIndex        =   20
         Top             =   1230
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         ListField       =   "banc_tx_descricao"
         BoundColumn     =   "banc_cd_codigo"
         Text            =   ""
      End
      Begin VB.Label LblVerso 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Verso:"
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
         Height          =   285
         Left            =   150
         TabIndex        =   39
         Top             =   1620
         Width           =   1305
      End
      Begin VB.Label LblNroCheque 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº do Cheque:"
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
         Left            =   135
         TabIndex        =   38
         Top             =   548
         Width           =   1320
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   105
         TabIndex        =   37
         Top             =   915
         Width           =   1350
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal:"
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
         Left            =   150
         TabIndex        =   36
         Top             =   1275
         Width           =   1305
      End
   End
   Begin Threed.SSPanel PanDetalhesCartaTransferencia 
      Height          =   3300
      Left            =   720
      TabIndex        =   112
      Top             =   960
      Visible         =   0   'False
      Width           =   6930
      _Version        =   65536
      _ExtentX        =   12224
      _ExtentY        =   5821
      _StockProps     =   15
      Caption         =   "Detalhes do Recibo"
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
      Begin VB.TextBox txtNumeroRecibo 
         Height          =   285
         Left            =   1320
         TabIndex        =   119
         Top             =   1080
         Width           =   3015
      End
      Begin VB.OptionButton optModelo3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modelo 3"
         Height          =   375
         Left            =   5040
         TabIndex        =   118
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optModelo2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modelo 2"
         Height          =   375
         Left            =   2640
         TabIndex        =   117
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optModelo1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modelo 1"
         Height          =   375
         Left            =   240
         TabIndex        =   116
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox txtEmissor 
         Height          =   315
         Left            =   1320
         TabIndex        =   121
         Top             =   2160
         Visible         =   0   'False
         Width           =   5430
      End
      Begin VB.CommandButton cmdConfirmarCartaTransferencia 
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
         Left            =   4200
         TabIndex        =   122
         Top             =   2880
         Width           =   1140
      End
      Begin VB.CommandButton cmdRetornarCartaTransfencia 
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
         Left            =   5520
         TabIndex        =   123
         Top             =   2880
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker DtpDataCartaTransferencia 
         Height          =   315
         Left            =   1320
         TabIndex        =   120
         Top             =   1440
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         DateIsNull      =   -1  'True
         Format          =   81657857
         CurrentDate     =   37460
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nº Recibo:"
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
         TabIndex        =   115
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Data:"
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
         TabIndex        =   114
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label LbLEmissor 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Emissor:"
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
         TabIndex        =   113
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin Threed.SSPanel PanTextoRecibo 
      Height          =   2220
      Left            =   5040
      TabIndex        =   54
      Top             =   360
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
      Alignment       =   8
      Begin VB.TextBox TxtTextoRecibo 
         Height          =   1335
         Left            =   150
         MultiLine       =   -1  'True
         TabIndex        =   57
         Text            =   "TelaDuplicatas.frx":02BF
         Top             =   300
         Width           =   6630
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
         TabIndex        =   56
         Top             =   1785
         Width           =   1140
      End
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
         TabIndex        =   55
         Top             =   1785
         Width           =   1140
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7560
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSPanel PanBotoes 
      Height          =   1500
      Left            =   0
      TabIndex        =   15
      Top             =   4440
      Width           =   12735
      _Version        =   65536
      _ExtentX        =   22463
      _ExtentY        =   2646
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
      Begin VB.CheckBox chkTotalFiltroDesconto 
         Caption         =   "Somar Descontos/Acréscimos"
         Height          =   255
         Left            =   8400
         TabIndex        =   125
         Top             =   0
         Width           =   2535
      End
      Begin VB.CommandButton cmdContabiliza 
         Caption         =   "Contabilizar"
         Height          =   315
         Left            =   9240
         TabIndex        =   124
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdCartaTransferencia 
         BackColor       =   &H00000000&
         Caption         =   "Car&ta de Transf."
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
         Height          =   300
         Left            =   8760
         TabIndex        =   111
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdExport 
         Height          =   315
         Left            =   8760
         Picture         =   "TelaDuplicatas.frx":02C5
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   1080
         Width           =   375
      End
      Begin VB.CheckBox chkTotalFiltroReceber 
         Caption         =   "Somar Valores a Receber"
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   0
         Width           =   2295
      End
      Begin VB.CheckBox chkTotalFiltroPagar 
         Caption         =   "Somar Valores a Pagar"
         Height          =   255
         Left            =   4320
         TabIndex        =   106
         Top             =   0
         Width           =   1935
      End
      Begin VB.CommandButton cmdCancelarDuplicata 
         Caption         =   "Cancelar"
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
         Left            =   8760
         TabIndex        =   100
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton CmdAltParcela 
         BackColor       =   &H00000000&
         Caption         =   "&Alt. Parcela"
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
         Left            =   2520
         TabIndex        =   6
         Top             =   720
         Width           =   1475
      End
      Begin VB.CommandButton CmdCanExpBanco 
         Caption         =   "E&xp. Banco"
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
         Left            =   7200
         TabIndex        =   9
         Top             =   720
         Width           =   1475
      End
      Begin VB.CommandButton CmdRecibo 
         BackColor       =   &H00000000&
         Caption         =   "Rec&ibo"
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
         Height          =   300
         Left            =   5640
         TabIndex        =   8
         Top             =   720
         Width           =   1475
      End
      Begin VB.CommandButton CmdCopia 
         Caption         =   "&Cópia Cheque"
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
         Left            =   2520
         TabIndex        =   1
         Top             =   360
         Width           =   1475
      End
      Begin VB.CommandButton CmdProcPagto 
         Caption         =   "P&roc. Pagto."
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
         Left            =   960
         TabIndex        =   10
         Top             =   1080
         Width           =   1475
      End
      Begin VB.CommandButton CmdEstorno 
         BackColor       =   &H00000000&
         Caption         =   "Es&torno"
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
         Left            =   7200
         TabIndex        =   4
         Top             =   360
         Width           =   1475
      End
      Begin VB.CommandButton CmdPagamento 
         BackColor       =   &H00000000&
         Caption         =   "&Efetivar"
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
         Left            =   960
         TabIndex        =   0
         Top             =   360
         Width           =   1475
      End
      Begin VB.CommandButton CmdDescAcresc 
         BackColor       =   &H00000000&
         Caption         =   "&Desc/Acrésc"
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
         Left            =   960
         TabIndex        =   5
         Top             =   720
         Width           =   1475
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Pe&ndente"
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
         Left            =   5640
         TabIndex        =   3
         Top             =   360
         Width           =   1475
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
         Height          =   300
         Left            =   2520
         TabIndex        =   11
         Top             =   1080
         Width           =   1475
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
         Height          =   300
         Left            =   4080
         TabIndex        =   12
         Top             =   1080
         Width           =   1475
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
         Height          =   300
         Left            =   5640
         TabIndex        =   13
         Top             =   1080
         Width           =   1475
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
         Height          =   300
         Left            =   7200
         TabIndex        =   14
         Top             =   1080
         Width           =   1475
      End
      Begin VB.CommandButton CmdBaixaParcial 
         BackColor       =   &H00000000&
         Caption         =   "&BaixaParcial"
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
         Left            =   4080
         TabIndex        =   2
         Top             =   360
         Width           =   1475
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
         Height          =   300
         Left            =   4080
         TabIndex        =   7
         Top             =   720
         Width           =   1475
      End
      Begin VB.Label lblTotalFiltro3 
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
         Left            =   10920
         TabIndex        =   126
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblTotalFiltro2 
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
         Left            =   2400
         TabIndex        =   108
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblTotalFiltro1 
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
         Left            =   6360
         TabIndex        =   107
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1125
      Top             =   3525
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
   Begin MSDataListLib.DataCombo cboLocalizar 
      Bindings        =   "TelaDuplicatas.frx":0474
      Height          =   315
      Left            =   2400
      TabIndex        =   32
      Top             =   3180
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "dupl_dt_vencimento"
      BoundColumn     =   "dupl_cd_duplicata"
      Text            =   ""
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Bindings        =   "TelaDuplicatas.frx":0489
      Height          =   3120
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   5503
      _LayoutType     =   4
      _RowHeight      =   25
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Vencimento"
      Columns(0).DataField=   "dupl_dt_vencimento"
      Columns(0).NumberFormat=   "dd/mm/yy"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nº Proc. Pag./Rec."
      Columns(1).DataField=   "Parcela"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nº Doc."
      Columns(2).DataField=   "Numdocumento"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "A Receber"
      Columns(3).DataField=   "entrada"
      Columns(3).NumberFormat=   "Standard"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "A Pagar"
      Columns(4).DataField=   "saida"
      Columns(4).NumberFormat=   "Standard"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Desconto/Acréscimo"
      Columns(5).DataField=   "DA"
      Columns(5).NumberFormat=   "Standard"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Pagto./Rec."
      Columns(6).DataField=   "dupl_dt_pagrec"
      Columns(6).NumberFormat=   "dd/mm/yy"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Status"
      Columns(7).DataField=   "status"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Competência"
      Columns(8).DataField=   "Competencia"
      Columns(8).NumberFormat=   "dd/mm/yy"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Cód Reduzido"
      Columns(9).DataField=   "CodReduzido"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Cliente/Fornecedor"
      Columns(10).DataField=   "focl_tx_fantasia"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Tipo Doc."
      Columns(11).DataField=   "TipoDocDesc"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Processo Pagamento"
      Columns(12).DataField=   "dupl_dt_procpagamento"
      Columns(12).NumberFormat=   "dd/mm/yy"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Data Conciliação"
      Columns(13).DataField=   "dupl_dt_concilia"
      Columns(13).NumberFormat=   "dd/mm/yy"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Tipo Pagto./Rec."
      Columns(14).DataField=   "tipopagto"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "Nº Doc. Pagto./Rec."
      Columns(15).DataField=   "dupl_tx_docpagrec"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "Conta Corrente (Empresa)"
      Columns(16).DataField=   "baempresa"
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "Conta Corrente (Fornecedor)"
      Columns(17).DataField=   "bafornecedor"
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "Forma Pagto"
      Columns(18).DataField=   "fopa_tx_descricao"
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "Exp. Cont."
      Columns(19).DataField=   "exporta"
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   16
      Columns(20)._MaxComboItems=   5
      Columns(20).ValueItems(0)._DefaultItem=   0
      Columns(20).ValueItems(0).Value=   "S"
      Columns(20).ValueItems(0).Value.vt=   8
      Columns(20).ValueItems(0).DisplayValue=   "Sim"
      Columns(20).ValueItems(0).DisplayValue.vt=   8
      Columns(20).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(20).ValueItems.Count=   1
      Columns(20).Caption=   "Estorno"
      Columns(20).DataField=   "dupl_tx_Estorno"
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "Codigo"
      Columns(21).DataField=   "dupl_cd_duplicata"
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "fatura"
      Columns(22).DataField=   "dupl_nr_fatura"
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   16
      Columns(23)._MaxComboItems=   5
      Columns(23).ValueItems(0)._DefaultItem=   0
      Columns(23).ValueItems(0).Value=   "N"
      Columns(23).ValueItems(0).Value.vt=   8
      Columns(23).ValueItems(0).DisplayValue=   "Normal"
      Columns(23).ValueItems(0).DisplayValue.vt=   8
      Columns(23).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(23).ValueItems(1)._DefaultItem=   0
      Columns(23).ValueItems(1).Value=   "P"
      Columns(23).ValueItems(1).Value.vt=   8
      Columns(23).ValueItems(1).DisplayValue=   "Pendente"
      Columns(23).ValueItems(1).DisplayValue.vt=   8
      Columns(23).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(23).ValueItems.Count=   2
      Columns(23).Caption=   "Estoque"
      Columns(23).DataField=   "Estoque"
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "Exp. Banco"
      Columns(24).DataField=   "ExpBanco"
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   25
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=25"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2011"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=529"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(0)._MinWidth=64"
      Splits(0)._ColumnProps(10)=   "Column(1).Width=1614"
      Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=1535"
      Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=529"
      Splits(0)._ColumnProps(15)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(16)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(17)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(18)=   "Column(1)._MinWidth=56372864"
      Splits(0)._ColumnProps(19)=   "Column(2).Width=1773"
      Splits(0)._ColumnProps(20)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(2)._WidthInPix=1693"
      Splits(0)._ColumnProps(22)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(2)._ColStyle=532"
      Splits(0)._ColumnProps(24)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(25)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(26)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(27)=   "Column(2)._MinWidth=55182876"
      Splits(0)._ColumnProps(28)=   "Column(3).Width=1773"
      Splits(0)._ColumnProps(29)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(3)._WidthInPix=1693"
      Splits(0)._ColumnProps(31)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(32)=   "Column(3)._ColStyle=530"
      Splits(0)._ColumnProps(33)=   "Column(3).WrapText=1"
      Splits(0)._ColumnProps(34)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(35)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(36)=   "Column(3)._MinWidth=55182876"
      Splits(0)._ColumnProps(37)=   "Column(4).Width=1693"
      Splits(0)._ColumnProps(38)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(4)._WidthInPix=1614"
      Splits(0)._ColumnProps(40)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(41)=   "Column(4)._ColStyle=530"
      Splits(0)._ColumnProps(42)=   "Column(4).WrapText=1"
      Splits(0)._ColumnProps(43)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(44)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(45)=   "Column(4)._MinWidth=24637548"
      Splits(0)._ColumnProps(46)=   "Column(5).Width=1799"
      Splits(0)._ColumnProps(47)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(5)._WidthInPix=1720"
      Splits(0)._ColumnProps(49)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(50)=   "Column(5)._ColStyle=530"
      Splits(0)._ColumnProps(51)=   "Column(5).WrapText=1"
      Splits(0)._ColumnProps(52)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(53)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(54)=   "Column(6).Width=2037"
      Splits(0)._ColumnProps(55)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(6)._WidthInPix=1958"
      Splits(0)._ColumnProps(57)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(58)=   "Column(6)._ColStyle=529"
      Splits(0)._ColumnProps(59)=   "Column(6).WrapText=1"
      Splits(0)._ColumnProps(60)=   "Column(6).AllowFocus=0"
      Splits(0)._ColumnProps(61)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(62)=   "Column(7).Width=2355"
      Splits(0)._ColumnProps(63)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(64)=   "Column(7)._WidthInPix=2275"
      Splits(0)._ColumnProps(65)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(66)=   "Column(7)._ColStyle=528"
      Splits(0)._ColumnProps(67)=   "Column(7).WrapText=1"
      Splits(0)._ColumnProps(68)=   "Column(7).AllowFocus=0"
      Splits(0)._ColumnProps(69)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(70)=   "Column(7)._MinWidth=55229656"
      Splits(0)._ColumnProps(71)=   "Column(8).Width=2090"
      Splits(0)._ColumnProps(72)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(73)=   "Column(8)._WidthInPix=2011"
      Splits(0)._ColumnProps(74)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(75)=   "Column(8)._ColStyle=529"
      Splits(0)._ColumnProps(76)=   "Column(8).WrapText=1"
      Splits(0)._ColumnProps(77)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(78)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(79)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(80)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(81)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(82)=   "Column(9)._ColStyle=529"
      Splits(0)._ColumnProps(83)=   "Column(9).WrapText=1"
      Splits(0)._ColumnProps(84)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(85)=   "Column(10).Width=5106"
      Splits(0)._ColumnProps(86)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(87)=   "Column(10)._WidthInPix=5027"
      Splits(0)._ColumnProps(88)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(89)=   "Column(10)._ColStyle=528"
      Splits(0)._ColumnProps(90)=   "Column(10).WrapText=1"
      Splits(0)._ColumnProps(91)=   "Column(10).AllowFocus=0"
      Splits(0)._ColumnProps(92)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(93)=   "Column(10)._MinWidth=10"
      Splits(0)._ColumnProps(94)=   "Column(11).Width=2593"
      Splits(0)._ColumnProps(95)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(96)=   "Column(11)._WidthInPix=2514"
      Splits(0)._ColumnProps(97)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(98)=   "Column(11)._ColStyle=528"
      Splits(0)._ColumnProps(99)=   "Column(11).WrapText=1"
      Splits(0)._ColumnProps(100)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(101)=   "Column(12).Width=1905"
      Splits(0)._ColumnProps(102)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(103)=   "Column(12)._WidthInPix=1826"
      Splits(0)._ColumnProps(104)=   "Column(12)._EditAlways=0"
      Splits(0)._ColumnProps(105)=   "Column(12)._ColStyle=529"
      Splits(0)._ColumnProps(106)=   "Column(12).WrapText=1"
      Splits(0)._ColumnProps(107)=   "Column(12).AllowFocus=0"
      Splits(0)._ColumnProps(108)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(109)=   "Column(13).Width=2011"
      Splits(0)._ColumnProps(110)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(111)=   "Column(13)._WidthInPix=1931"
      Splits(0)._ColumnProps(112)=   "Column(13)._EditAlways=0"
      Splits(0)._ColumnProps(113)=   "Column(13)._ColStyle=529"
      Splits(0)._ColumnProps(114)=   "Column(13).WrapText=1"
      Splits(0)._ColumnProps(115)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(116)=   "Column(14).Width=2064"
      Splits(0)._ColumnProps(117)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(118)=   "Column(14)._WidthInPix=1984"
      Splits(0)._ColumnProps(119)=   "Column(14)._EditAlways=0"
      Splits(0)._ColumnProps(120)=   "Column(14)._ColStyle=528"
      Splits(0)._ColumnProps(121)=   "Column(14).WrapText=1"
      Splits(0)._ColumnProps(122)=   "Column(14).AllowFocus=0"
      Splits(0)._ColumnProps(123)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(124)=   "Column(14)._MinWidth=8"
      Splits(0)._ColumnProps(125)=   "Column(15).Width=2223"
      Splits(0)._ColumnProps(126)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(127)=   "Column(15)._WidthInPix=2143"
      Splits(0)._ColumnProps(128)=   "Column(15)._EditAlways=0"
      Splits(0)._ColumnProps(129)=   "Column(15)._ColStyle=528"
      Splits(0)._ColumnProps(130)=   "Column(15).WrapText=1"
      Splits(0)._ColumnProps(131)=   "Column(15).AllowFocus=0"
      Splits(0)._ColumnProps(132)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(133)=   "Column(16).Width=3096"
      Splits(0)._ColumnProps(134)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(135)=   "Column(16)._WidthInPix=3016"
      Splits(0)._ColumnProps(136)=   "Column(16)._EditAlways=0"
      Splits(0)._ColumnProps(137)=   "Column(16)._ColStyle=528"
      Splits(0)._ColumnProps(138)=   "Column(16).WrapText=1"
      Splits(0)._ColumnProps(139)=   "Column(16).AllowFocus=0"
      Splits(0)._ColumnProps(140)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(141)=   "Column(17).Width=2858"
      Splits(0)._ColumnProps(142)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(143)=   "Column(17)._WidthInPix=2778"
      Splits(0)._ColumnProps(144)=   "Column(17)._EditAlways=0"
      Splits(0)._ColumnProps(145)=   "Column(17)._ColStyle=529"
      Splits(0)._ColumnProps(146)=   "Column(17).WrapText=1"
      Splits(0)._ColumnProps(147)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(148)=   "Column(18).Width=2725"
      Splits(0)._ColumnProps(149)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(150)=   "Column(18)._WidthInPix=2646"
      Splits(0)._ColumnProps(151)=   "Column(18)._EditAlways=0"
      Splits(0)._ColumnProps(152)=   "Column(18)._ColStyle=528"
      Splits(0)._ColumnProps(153)=   "Column(18).WrapText=1"
      Splits(0)._ColumnProps(154)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(155)=   "Column(19).Width=1429"
      Splits(0)._ColumnProps(156)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(157)=   "Column(19)._WidthInPix=1349"
      Splits(0)._ColumnProps(158)=   "Column(19)._EditAlways=0"
      Splits(0)._ColumnProps(159)=   "Column(19)._ColStyle=528"
      Splits(0)._ColumnProps(160)=   "Column(19).WrapText=1"
      Splits(0)._ColumnProps(161)=   "Column(19).AllowFocus=0"
      Splits(0)._ColumnProps(162)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(163)=   "Column(20).Width=1402"
      Splits(0)._ColumnProps(164)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(165)=   "Column(20)._WidthInPix=1323"
      Splits(0)._ColumnProps(166)=   "Column(20)._EditAlways=0"
      Splits(0)._ColumnProps(167)=   "Column(20)._ColStyle=528"
      Splits(0)._ColumnProps(168)=   "Column(20).WrapText=1"
      Splits(0)._ColumnProps(169)=   "Column(20).AllowFocus=0"
      Splits(0)._ColumnProps(170)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(171)=   "Column(21).Width=1244"
      Splits(0)._ColumnProps(172)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(173)=   "Column(21)._WidthInPix=1164"
      Splits(0)._ColumnProps(174)=   "Column(21)._EditAlways=0"
      Splits(0)._ColumnProps(175)=   "Column(21)._ColStyle=532"
      Splits(0)._ColumnProps(176)=   "Column(21).Visible=0"
      Splits(0)._ColumnProps(177)=   "Column(21).WrapText=1"
      Splits(0)._ColumnProps(178)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(179)=   "Column(22).Width=2725"
      Splits(0)._ColumnProps(180)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(181)=   "Column(22)._WidthInPix=2646"
      Splits(0)._ColumnProps(182)=   "Column(22)._EditAlways=0"
      Splits(0)._ColumnProps(183)=   "Column(22)._ColStyle=532"
      Splits(0)._ColumnProps(184)=   "Column(22).Visible=0"
      Splits(0)._ColumnProps(185)=   "Column(22).WrapText=1"
      Splits(0)._ColumnProps(186)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(187)=   "Column(23).Width=1482"
      Splits(0)._ColumnProps(188)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(189)=   "Column(23)._WidthInPix=1402"
      Splits(0)._ColumnProps(190)=   "Column(23)._EditAlways=0"
      Splits(0)._ColumnProps(191)=   "Column(23)._ColStyle=528"
      Splits(0)._ColumnProps(192)=   "Column(23).Visible=0"
      Splits(0)._ColumnProps(193)=   "Column(23).WrapText=1"
      Splits(0)._ColumnProps(194)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(195)=   "Column(24).Width=2725"
      Splits(0)._ColumnProps(196)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(197)=   "Column(24)._WidthInPix=2646"
      Splits(0)._ColumnProps(198)=   "Column(24)._EditAlways=0"
      Splits(0)._ColumnProps(199)=   "Column(24)._ColStyle=532"
      Splits(0)._ColumnProps(200)=   "Column(24).WrapText=1"
      Splits(0)._ColumnProps(201)=   "Column(24).Order=25"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=104,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=111,.parent=1,.valignment=2,.wraptext=-1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=120,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=112,.parent=2,.alignment=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=113,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=114,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=116,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=115,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=117,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=118,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=119,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=121,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=122,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=126,.parent=111,.alignment=2"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=123,.parent=112"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=124,.parent=113"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=125,.parent=115"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=130,.parent=111,.alignment=2"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=127,.parent=112,.alignment=2"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=128,.parent=113"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=129,.parent=115"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=111"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=112"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=113"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=115"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=134,.parent=111,.alignment=1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=131,.parent=112"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=132,.parent=113"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=133,.parent=115"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=138,.parent=111,.alignment=1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=135,.parent=112"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=136,.parent=113"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=137,.parent=115"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=142,.parent=111,.alignment=1"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=139,.parent=112,.alignment=2"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=140,.parent=113"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=141,.parent=115"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=146,.parent=111,.alignment=2"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=143,.parent=112,.alignment=2"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=144,.parent=113"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=145,.parent=115"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=150,.parent=111,.alignment=0"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=147,.parent=112,.alignment=2"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=148,.parent=113"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=149,.parent=115"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=50,.parent=111,.alignment=2"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=47,.parent=112"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=48,.parent=113"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=49,.parent=115"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=58,.parent=111,.alignment=2"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=55,.parent=112"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=56,.parent=113"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=57,.parent=115"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=154,.parent=111,.alignment=0"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=151,.parent=112,.alignment=2"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=152,.parent=113"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=153,.parent=115"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=24,.parent=111,.alignment=0"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=21,.parent=112,.alignment=2"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=22,.parent=113"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=23,.parent=115"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=162,.parent=111,.alignment=2"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=159,.parent=112"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=160,.parent=113"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=161,.parent=115"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=20,.parent=111,.alignment=2"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=17,.parent=112,.alignment=2"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=18,.parent=113"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=19,.parent=115"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=166,.parent=111,.alignment=0"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=163,.parent=112,.alignment=2"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=164,.parent=113"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=165,.parent=115"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=170,.parent=111,.alignment=0"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=167,.parent=112,.alignment=2"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=168,.parent=113"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=169,.parent=115"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=174,.parent=111,.alignment=0"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=171,.parent=112,.alignment=2"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=172,.parent=113"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=173,.parent=115"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=16,.parent=111,.alignment=2"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=13,.parent=112,.alignment=2"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=14,.parent=113"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=15,.parent=115"
      _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=54,.parent=111,.alignment=0"
      _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=51,.parent=112"
      _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=52,.parent=113"
      _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=53,.parent=115"
      _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=178,.parent=111,.alignment=0"
      _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=175,.parent=112,.alignment=2"
      _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=176,.parent=113"
      _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=177,.parent=115"
      _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=182,.parent=111,.alignment=0"
      _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=179,.parent=112,.alignment=2"
      _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=180,.parent=113"
      _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=181,.parent=115"
      _StyleDefs(120) =   "Splits(0).Columns(21).Style:id=190,.parent=111"
      _StyleDefs(121) =   "Splits(0).Columns(21).HeadingStyle:id=187,.parent=112"
      _StyleDefs(122) =   "Splits(0).Columns(21).FooterStyle:id=188,.parent=113"
      _StyleDefs(123) =   "Splits(0).Columns(21).EditorStyle:id=189,.parent=115"
      _StyleDefs(124) =   "Splits(0).Columns(22).Style:id=194,.parent=111"
      _StyleDefs(125) =   "Splits(0).Columns(22).HeadingStyle:id=191,.parent=112"
      _StyleDefs(126) =   "Splits(0).Columns(22).FooterStyle:id=192,.parent=113"
      _StyleDefs(127) =   "Splits(0).Columns(22).EditorStyle:id=193,.parent=115"
      _StyleDefs(128) =   "Splits(0).Columns(23).Style:id=32,.parent=111,.alignment=0"
      _StyleDefs(129) =   "Splits(0).Columns(23).HeadingStyle:id=29,.parent=112,.alignment=2"
      _StyleDefs(130) =   "Splits(0).Columns(23).FooterStyle:id=30,.parent=113"
      _StyleDefs(131) =   "Splits(0).Columns(23).EditorStyle:id=31,.parent=115"
      _StyleDefs(132) =   "Splits(0).Columns(24).Style:id=46,.parent=111"
      _StyleDefs(133) =   "Splits(0).Columns(24).HeadingStyle:id=43,.parent=112"
      _StyleDefs(134) =   "Splits(0).Columns(24).FooterStyle:id=44,.parent=113"
      _StyleDefs(135) =   "Splits(0).Columns(24).EditorStyle:id=45,.parent=115"
      _StyleDefs(136) =   "Named:id=33:Normal"
      _StyleDefs(137) =   ":id=33,.parent=0"
      _StyleDefs(138) =   "Named:id=34:Heading"
      _StyleDefs(139) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(140) =   ":id=34,.wraptext=-1"
      _StyleDefs(141) =   "Named:id=35:Footing"
      _StyleDefs(142) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(143) =   "Named:id=36:Selected"
      _StyleDefs(144) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(145) =   "Named:id=37:Caption"
      _StyleDefs(146) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(147) =   "Named:id=38:HighlightRow"
      _StyleDefs(148) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(149) =   "Named:id=39:EvenRow"
      _StyleDefs(150) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(151) =   "Named:id=40:OddRow"
      _StyleDefs(152) =   ":id=40,.parent=33"
      _StyleDefs(153) =   "Named:id=41:RecordSelector"
      _StyleDefs(154) =   ":id=41,.parent=34"
      _StyleDefs(155) =   "Named:id=42:FilterBar"
      _StyleDefs(156) =   ":id=42,.parent=33"
   End
   Begin VB.Label lblLocalizar 
      Alignment       =   1  'Right Justify
      Caption         =   "Localizar Vencimento:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   75
      TabIndex        =   41
      Top             =   3225
      Width           =   2295
   End
End
Attribute VB_Name = "TelaDuplicatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Contador As Double 'Conta o número de saldos calculados
Dim VetorSaldo() As Single 'Esse vetor irá conter o saldo de cada linha e coluna
Dim Filtro As Boolean 'Verifica se existe filtro, caso exista é dado um refresh no controle para atualizar os dados
Dim XFT_DOCPAGREC As String
Dim XFT_TEXTO As String     'Texto do recibo
Dim Posicao As Variant

'Variáveis para a impressão do cheque
Dim XNumCheque As String
Dim XDtPagamento As Date
Dim xnota As String, Xhistorico As String
Dim XLI_TIPODOC As Integer

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
Dim Modelo As Integer ' Indica qual o tipo de recibo a ser impresso

Function FunVerificacoes() As Boolean
    
    Dim XLT_SQL As String
    Dim XLO_CHECK As ADODB.Recordset
    'Não considerar quando tiver desconto/acréscimo pois o cancelamento
    'é feito pelo valor bruto da nota que é a provisão
    
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
    XLT_SQL = "SELECT tido_tx_Nec FROM tipodocumentos WHERE tido_tx_descricao = '" & Adodc1.Recordset.Fields("tipodocdesc") & "'"
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
'                    "WHERE (lanc_tx_Origem = 'FN' or lanc_tx_Origem = 'T') AND nofi_nr_fatura = " & Adodc1.Recordset.Fields("dupl_nr_fatura") & ""
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
Sub subCompromissarCancelamento()
    
    Dim XLD_DATA As Date 'Data de cancelamento da duplicata
    Dim XLT_SQL As String
    Dim XLO_CHECK As ADODB.Recordset
    
    XLT_SQL = "SELECT lanc_dt_data FROM Lancamentos " & _
                " WHERE lanc_tx_Origem = 'FD' AND " & _
                " empr_cd_empresa = " & PCodEmpresa & " AND " & _
                " lanc_cd_OrigemFin = " & Adodc1.Recordset.Fields("dupl_cd_duplicata") & ""
    SubQOpenRecordset XLO_CHECK, XLT_SQL, Estatico
    
    If Not XLO_CHECK.EOF Then
        'verificar se o mês está fechado
        XLD_DATA = XLO_CHECK!lanc_dt_Data
        XLO_CHECK.Close
    Else
        'se não encontrar nenhum registro significa que não existe nenhuma duplicata cancelada na contabilidade para esse registro
        MsgBox "Não foi possível executar essa operação. Entre em contato com o suporte!", vbCritical, "ANTENÇÃO"
        XLO_CHECK.Close
        Exit Sub
    End If
    
    'Verifica se o mês está fechado
    If funMesFechado(Format(XLD_DATA, "dd/MM/yyyy")) <> "A" Then
        MsgBox "Não é possível realizar a operação nessa data, pois o mês contábil já foi fechado. " _
            + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
         Me.panCancelar.Visible = False
        Exit Sub
    End If
    Conexao.BeginTrans
    
    'excluindo o registro da contabilidade
    Conexao.Execute "DELETE FROM lancamentos WHERE lanc_tx_Origem = 'FD' AND " & _
                " empr_cd_empresa = " & PCodEmpresa & " AND " & _
                " lanc_cd_OrigemFin = " & Adodc1.Recordset.Fields("dupl_cd_duplicata") & ""
    
    'atualizando a duplicata
    Conexao.Execute "UPDATE duplicatas SET dupl_tx_status = 'C' WHERE " & _
                    " dupl_cd_duplicata = " & Adodc1.Recordset.Fields("dupl_cd_duplicata") & ""
    
    'atualizando a nota fiscal
    Conexao.Execute "UPDATE NotasFiscais SET nofi_tx_status = 'C' WHERE " & _
                    " empr_cd_empresa = " & PCodEmpresa & " AND " & _
                    " nofi_nr_fatura = " & Adodc1.Recordset.Fields("dupl_nr_fatura") & ""
    
    Conexao.CommitTrans
    
    'recarrega o grid
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    
    'emite a mensagem
    MsgBox "Operação realizada com sucesso!", vbInformation, "Cancelamento"
    
End Sub
Sub subImprimeRecibo(NumRecibo As String) 'Alteração em 24/09/08 - Patrícia.Anterior: Sub subImprimeRecibo(NumRecibo As Double)
    Dim XCidade As String
    Dim Xhistorico As String
    Dim XExtenso As String
    Dim XLT_TEMDESC As String
    Dim Xvalor As Double
    Dim XLF_VLBRUTO As Double
    Dim XLF_VALORGLOSA As Double
    Dim XLO_SUBRELATORIO As ADODB.Recordset
    Dim XLO_SUBRELATORIO2 As ADODB.Recordset
    Dim XLO_SUBRELATORIO3 As ADODB.Recordset
    Dim XLO_NOTA As ADODB.Recordset
    Dim XLO_RSEMPRESA As ADODB.Recordset ' Declaração em 24/09/08 - Patrícia
    Dim XLB_FORNCLIENTE As Boolean
        
    TelaDuplicatas.MousePointer = vbHourglass
    
    SubQOpenRecordset XLO_NOTA, "SELECT nofi_tx_Historico,focl_cd_Forncli FROM NotasFiscais WHERE nofi_nr_Fatura = " & Adodc1.Recordset.Fields("dupl_nr_Fatura"), Estatico
    Xhistorico = XLO_NOTA!nofi_tx_historico
    If Not IsNull(XLO_NOTA!focl_cd_forncli) Then ' 03/10/2008 - Patrícia
        Set XLO_SUBRELATORIO3 = ConexaoRelatorio.Execute("SELECT * FROM FornClientes WHERE focl_cd_Forncli = " & XLO_NOTA!focl_cd_forncli & "")
        XLB_FORNCLIENTE = True
    Else
        XLB_FORNCLIENTE = False
    End If
    XLO_NOTA.Close
    
    XCidade = "Salvador"
    XLT_TEMDESC = "N"
  
    If Adodc1.Recordset.Fields("dupl_tx_DebCred") = "D" Then
        Xvalor = CStr(TDBGrid1.Columns("A Pagar"))
        XLF_VLBRUTO = CStr(TDBGrid1.Columns("A Pagar"))
        XExtenso = funExtenso(TDBGrid1.Columns("A Pagar"))
        FrmVisRelatoriosEXT.Caption = "Recibo (Contas a Pagar)"
    Else
        Xvalor = CStr(TDBGrid1.Columns("A Receber"))
        XLF_VLBRUTO = CStr(TDBGrid1.Columns("A Receber"))
        XExtenso = funExtenso(TDBGrid1.Columns("A Receber"))
        FrmVisRelatoriosEXT.Caption = "Recibo (Contas a Receber)"
    End If
        
    'DESC/ACRES
    If TDBGrid1.Columns("Desconto/Acréscimo") <> "0,00" Then XLT_TEMDESC = "S"
    Xvalor = CStr(Format(CDbl(XLF_VLBRUTO) + TDBGrid1.Columns("Desconto/Acréscimo"), "standard"))
    XExtenso = funExtenso(CDbl(Xvalor))

    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDuplicatas WHERE dupl_cd_duplicata = " & Adodc1.Recordset.Fields("dupl_cd_duplicata") & "")
    Set XLO_SUBRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINRelFinanceiroDescAcresc WHERE dupl_nr_Fatura = " & Adodc1.Recordset.Fields("dupl_nr_fatura") & "")
    Set XLO_SUBRELATORIO2 = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDistribCopiaCheque WHERE nofi_nr_Fatura = " & Adodc1.Recordset.Fields("dupl_nr_fatura") & "")
    Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas")
    
    'Abre o relatório
    SubQOpenRecordset XLO_RSEMPRESA, "SELECT empr_tx_cliente, empr_tx_RazaoSocial, empr_tx_CGC FROM Empresas", "Estatico"
    If Not (XLO_RSEMPRESA.EOF) Then ' If inserido em 24/09/08 - Patrícia. Inclusão de relatórios solicitados pelo cliente Instituto Aliança
       If Trim(XLO_RSEMPRESA!empr_tx_cliente) = "ALIANCA" Then
            Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\ReciboDuplicata_Alianca.rpt")
            With XGR_RELATORIO.FormulaFields 'Inserido em 11/10/08 - Patrícia
                    .GetItemByName("RAZAOSOCIAL_EMPRESA").Text = Chr(34) + CStr(XLO_RSEMPRESA!empr_tx_RazaoSocial) + Chr(34)
                    .GetItemByName("CNPJ_EMPRESA").Text = Chr(34) + "CNPJ: " + CStr(XLO_RSEMPRESA!empr_tx_CGC) + Chr(34)
            End With

       Else
            Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\ReciboDuplicata.rpt")
       End If
    End If
    XLO_RSEMPRESA.Close

    
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO

    'Ativa os Relatórios
    With XGR_RELATORIO.FormulaFields
        .GetItemByName("Recibo").Text = Chr(34) + CStr(NumRecibo) + Chr(34)
        .GetItemByName("Empresa").Text = Chr(34) + CStr(PCodEmpresa) + Chr(34)
        .GetItemByName("Texto").Text = Chr(34) + FunTiraAspas(XFT_TEXTO) + Chr(34)
        .GetItemByName("NomeEmpresa").Text = Chr(34) + CStr(PEmpresa) + Chr(34)
        .GetItemByName("Valor").Text = Chr(34) + CStr(Format(Xvalor, "standard")) + Chr(34)
        .GetItemByName("ValorExtenso").Text = Chr(34) + XExtenso + Chr(34)
        .GetItemByName("Cidade").Text = Chr(34) + XCidade + Chr(34)
        .GetItemByName("TemDescAcresc").Text = Chr(34) + XLT_TEMDESC + Chr(34)
        .GetItemByName("Historico").Text = Chr(34) + FunTiraAspas(Xhistorico) + Chr(34)

    End With
    
    FrmVisRelatoriosEXT.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
    FrmVisRelatoriosEXT.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XLO_SUBRELATORIO, "Desconto"
    FrmVisRelatoriosEXT.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XLO_SUBRELATORIO2, "Distribuicao"
    If XLB_FORNCLIENTE = True Then ' 03/10/2008 - Patrícia
        FrmVisRelatoriosEXT.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XLO_SUBRELATORIO3, "Fornecedores"
    End If
    FrmVisRelatoriosEXT.SubVerRelatorio XGR_RELATORIO, ""
    
    TelaDuplicatas.MousePointer = vbDefault
End Sub


Sub subImprimeCopia()
  Dim total As String
  Dim Data As String
  Dim Xmes As String
  Dim Dia As Integer, Ano As Integer, Mes As Integer
  Dim xChave As Variant
  Dim XExtenso As String
  Dim XLA_SUBRELATORIO As ADODB.Recordset
  Dim XLO_SUBRELATORIO2 As ADODB.Recordset
  
  xChave = TDBGrid1.Bookmark
  TelaDuplicatas.MousePointer = vbHourglass
  
  If Adodc1.Recordset.Fields("da") <> "" Then
    total = CStr(Format((Adodc1.Recordset.Fields("valor") + Adodc1.Recordset.Fields("da")), "standard"))
  Else
    total = CStr(Format(Adodc1.Recordset.Fields("valor"), "standard"))
  End If
  Ano = Year(DtpDtPagto.Value)
  Mes = Month(DtpDtPagto.Value)
  Dia = Day(DtpDtPagto.Value)
  Xmes = MonthName(Mes, False)
  Data = (CStr(Dia) + " de " + Xmes + " de " + CStr(Ano))
  XExtenso = funExtenso(CDbl(total))
  
  'Abre os recordsets e verifica se eles possuem dados
  If CboCCorrente.Text <> "" Then
    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINCheques where (dupl_tx_debcred='D' and " & _
      "coco_cd_codigo = " & CboCCorrente.BoundText & " and dupl_tx_TipoPagRec = '1' and " & _
      "dupl_tx_docpagrec = '" & TxtNroCheque.Text & "' and (dupl_tx_Estorno<>'S' or (dupl_tx_estorno is null)))")
  Else ' BAIXA PARCIAL
    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINCheques where (dupl_tx_debcred='D' and " & _
      "coco_cd_codigo = " & CboCCorrenteParcial.BoundText & " and dupl_tx_TipoPagRec = '1' and " & _
      "dupl_tx_docpagrec = '" & TxtNroCheque.Text & "' and (dupl_tx_Estorno<>'S' or (dupl_tx_estorno is null)))")
  End If
  
  If XGO_RSRELATORIO.EOF Then
    MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
    Me.MousePointer = vbDefault
    Exit Sub
  End If
  
  Set XLA_SUBRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDistribCopiaCheque")
  Set XLO_SUBRELATORIO2 = ConexaoRelatorio.Execute("SELECT * FROM ConsFINCChequeDescAcresc")

  'Abre o relatório
  FrmVisRelatoriosEXT.Caption = "Cópia de Cheque"
  Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\CopiaCheque.rpt")
  
  'Passa o recordset para o relatório
  XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
  
  'Ativa os Relatórios
  With XGR_RELATORIO.FormulaFields
    .GetItemByName("Valor").Text = Chr(34) + Format(total, "standard") + Chr(34)
    .GetItemByName("Extenso").Text = Chr(34) + XExtenso + Chr(34)
    .GetItemByName("Cidade").Text = Chr(34) + TxtCidade.Text + Chr(34)
    .GetItemByName("NumDoc").Text = Chr(34) + TxtNroCheque.Text + Chr(34)
    .GetItemByName("Banco").Text = Chr(34) + CboCCorrente.Text + Chr(34)
    .GetItemByName("Historico").Text = Chr(34) + Xhistorico + Chr(34)
    .GetItemByName("NotaFiscal").Text = Chr(34) + xnota + Chr(34)
    .GetItemByName("Nominal").Text = Chr(34) + CboNominal.Text + Chr(34)
    .GetItemByName("Data").Text = Chr(34) + Data + Chr(34)
    .GetItemByName("Empresa").Text = Chr(34) + PEmpresa + Chr(34)
    .GetItemByName("NumFatura").Text = Chr(34) + TDBGrid1.Columns("Fatura") + Chr(34)
  End With
  FrmVisRelatoriosEXT.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XLA_SUBRELATORIO, "ConsDistribCopiaCheque"
  FrmVisRelatoriosEXT.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XLO_SUBRELATORIO2, "DescAcresc"
  FrmVisRelatoriosEXT.SubVerRelatorio XGR_RELATORIO, ""
  TelaDuplicatas.MousePointer = vbDefault
End Sub

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
    CmdPagamento.Enabled = True
    CmdDescAcresc.Enabled = True
    CmdAlterar.Enabled = True
    CmdCancelar.Enabled = True
    CmdSair.Enabled = True
    CmdImprimir.Enabled = True
    CmdOrdem.Enabled = True
    CmdEstorno.Enabled = True
    CmdProcPagto.Enabled = True
    CmdCopia.Enabled = True
    cboLocalizar.Enabled = True
    CmdCanExpBanco.Enabled = True
    CmdAltParcela.Enabled = True
    CmdBaixaParcial.Enabled = True
    'Verifica se tem Filtro Ativo
    If ContFil = 0 Then
       Botao_Fil.Enabled = False
       chkTotalFiltroPagar.Visible = False
       chkTotalFiltroPagar.Value = 0
       
       chkTotalFiltroReceber.Visible = False
       chkTotalFiltroReceber.Value = 0
       
       chkTotalFiltroDesconto.Visible = False
       chkTotalFiltroDesconto.Value = 0
       
       lblTotalFiltro1.Caption = ""
       lblTotalFiltro2.Caption = ""
       lblTotalFiltro3.Caption = ""
    Else
       Botao_Fil.Enabled = True
       
       chkTotalFiltroPagar.Visible = True
       chkTotalFiltroPagar.Value = 0
             
       chkTotalFiltroReceber.Visible = True
       chkTotalFiltroReceber.Value = 0
       
       chkTotalFiltroDesconto.Visible = True
       chkTotalFiltroDesconto.Value = 0
       
       lblTotalFiltro1.Caption = ""
       lblTotalFiltro2.Caption = ""
       lblTotalFiltro3.Caption = ""
        
    End If
End Sub

Private Sub subDesabilitaBotoes()
    CmdPagamento.Enabled = False
    CmdAlterar.Enabled = False
    CmdCancelar.Enabled = False
    Botao_Fil.Enabled = False
    CmdSair.Enabled = False
    CmdImprimir.Enabled = False
    CmdOrdem.Enabled = False
    CmdDescAcresc.Enabled = False
    CmdEstorno.Enabled = False
    CmdProcPagto.Enabled = False
    CmdCopia.Enabled = False
    CmdCanExpBanco.Enabled = False
    cboLocalizar.Enabled = False
    CmdAltParcela.Enabled = False
    CmdBaixaParcial.Enabled = False
End Sub

Private Sub Botao_Fil_Click()
    subDesabilitaBotoes
    
    Set Formulario = TelaDuplicatas
    
    subTelaValoresGlobais "P"
    TelaFiltro.Show 1 'Mostra a tela de filtros no modo modal
    subTelaValoresGlobais "G"
    
    subHabilitaBotoes

    If Filtrou = True Then
        subContaReg Adodc1
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
        Filtro = False
    End If
    
    If FunTabelaVazia(Adodc1) Then
       CmdAlterar.Enabled = False
       CmdCancelar.Enabled = False
       CmdPagamento.Enabled = False
       CmdDescAcresc.Enabled = False
       CmdOrdem.Enabled = False
       CmdImprimir.Enabled = False
    End If
    
    TDBGrid1.SetFocus
End Sub

Private Sub CboCCorrente_Change()
    If CboCCorrente.Text <> "" Then
        DatContaCorrente.Recordset.Bookmark = CboCCorrente.SelectedItem
        TxtCCorrente.Text = CboCCorrente.BoundText
        
        subConectarControleDadosNV DatFormaPagamento, "SELECT fopa_cd_codigo, banc_cd_codigo, CONVERT(VARCHAR, fopa_nr_numero) + ' - ' +  fopa_tx_descricao AS DescCombo FROM FormasPagamento WHERE banc_cd_codigo = " & DatContaCorrente.Recordset.Fields("banc_cd_codigo") & " ORDER BY fopa_nr_numero", Estatico
        
    End If
End Sub

Private Sub CboCCorrente2_Change()
    If CboCCorrente2.Text <> "" Then
        DatContaCorrente.Recordset.Bookmark = CboCCorrente2.SelectedItem
        TxtCCorrente2.Text = CboCCorrente2.BoundText
    End If
End Sub

Private Sub CboCCorrente3_Change()
    If CboCCorrente3.Text <> "" Then
        DatContaCorrente.Recordset.Bookmark = CboCCorrente3.SelectedItem
        TxtCCorrente3.Text = CboCCorrente3.BoundText
    End If
End Sub

Private Sub CboCCorrenteParcial_Change()
    If CboCCorrenteParcial.Text <> "" Then
        DatContaCorrente.Recordset.Bookmark = CboCCorrenteParcial.SelectedItem
        TxtCodCCorrenteParcial.Text = CboCCorrenteParcial.BoundText
    End If
End Sub

Private Sub CboCCorrenteResiduo_Change()
    If CboCCorrenteResiduo.Text <> "" Then
        DatContaCorrente.Recordset.Bookmark = CboCCorrenteResiduo.SelectedItem
        TxtCodCCorrenteResiduo.Text = CboCCorrenteResiduo.BoundText
    End If
End Sub


Private Sub cboLocalizar_Change()
    If cboLocalizar.BoundText <> "" Then
        Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
        
    End If
End Sub

Private Sub CboNominal_KeyPress(KeyAscii As Integer)
    If Len(CboNominal) = 50 Then
        If KeyAscii <> 8 And CboNominal.SelLength = 0 Then
            KeyAscii = 0
            Beep
        End If
    End If
End Sub


Private Sub CboTipoPag_LostFocus()
    If CboTipoPag.ItemData(CboTipoPag.ListIndex) = 6 Then
        TxtNumDoc.Text = ""
        TxtNumDoc.Enabled = False
    Else
        TxtNumDoc.Enabled = True
    End If
End Sub



Private Sub chkTotalFiltroPagar_Click()
 Dim total1 As Double
 Dim ResultString As String

    total1 = 0

    If (chkTotalFiltroPagar.Value = 1) Then
        
        lblTotalFiltro1.Visible = True
        
        TDBGrid1.MoveFirst
        
        While Not TDBGrid1.EOF
                    ResultString = TDBGrid1.Columns("A Pagar")
                    If ResultString = "" Then
                        total1 = total1 + 0
                        GoTo Proximo
                    End If
                    total1 = total1 + CDbl(TDBGrid1.Columns("A Pagar"))
Proximo:
                    TDBGrid1.MoveNext
        Wend
        
        lblTotalFiltro1.Caption = "R$ " & Format(CStr("R$ " & total1), "##,##0.00")
        
        If lblTotalFiltro1.Caption = "" Then
            lblTotalFiltro1.Caption = "R$0,00"
        End If
        
        TDBGrid1.MoveFirst
    Else
        lblTotalFiltro1.Caption = ""
    End If
    
End Sub



Private Sub chkTotalFiltroReceber_Click()

Dim total2 As Double
Dim ResultString As String
     
    total2 = 0

    If (chkTotalFiltroReceber.Value = 1) Then
        
        
        lblTotalFiltro2.Visible = True

        TDBGrid1.MoveFirst
        
        While Not TDBGrid1.EOF
                    ResultString = TDBGrid1.Columns("A Receber")
                    If ResultString = "" Then
                        total2 = total2 + 0
                        GoTo Proximo
                    End If
                        total2 = total2 + CDbl(TDBGrid1.Columns("A Receber"))
Proximo:
                    TDBGrid1.MoveNext
        Wend
        
        lblTotalFiltro2.Caption = "R$ " & Format(CStr("R$ " & total2), "##,##0.00")
        
        If lblTotalFiltro2.Caption = "" Then
            lblTotalFiltro2.Caption = "R$0,00"
        End If

        TDBGrid1.MoveFirst
    Else
        lblTotalFiltro2.Caption = ""

        
    End If
End Sub

Private Sub chkTotalFiltroDesconto_Click()

Dim total3 As Double
Dim ResultString As String
 
    total3 = 0
    
    If (chkTotalFiltroDesconto.Value = 1) Then
        
        
        lblTotalFiltro3.Visible = True

        TDBGrid1.MoveFirst
        
        While Not TDBGrid1.EOF
                    ResultString = TDBGrid1.Columns("Desconto/Acréscimo")
                    If ResultString = "" Then
                        total3 = total3 + 0
                        GoTo Proximo
                    End If
                        total3 = total3 + CDbl(TDBGrid1.Columns("Desconto/Acréscimo"))
Proximo:
                    TDBGrid1.MoveNext
        Wend
        
        lblTotalFiltro3.Caption = "R$ " & Format(CStr("R$ " & total3), "##,##0.00")
        
        If lblTotalFiltro3.Caption = "" Then
            lblTotalFiltro3.Caption = "R$0,00"
        End If
               

        TDBGrid1.MoveFirst
    Else
        lblTotalFiltro3.Caption = ""
      
    End If
    
End Sub

Private Sub CmdAlterar_Click()
    Dim ResFat As ADODB.Recordset
    Dim XTipoFatura As String
    Dim ResTipoDoc As ADODB.Recordset
    Dim xsigla As String
 

    XGB_EntrouDupl = True '27/05/10
    If Not FunVerifica_Permissao("TelaDuplicatas", "1") Then
        Exit Sub
    End If

    If Adodc1.Recordset.Fields("TipoDocDesc") = "GLOSA" Or Adodc1.Recordset.Fields("TipoDocDesc") = "NOTA FISCAL DE DEVOLUÇÃO" Then
        'Pegar nº do documento da nota de origem
        SubQOpenRecordset ResTipoDoc, "SELECT nofi_nr_documento FROM NotasFiscais WHERE nofi_cd_notafiscal = (SELECT nofi_cd_notaGlosaDev FROM NotasFiscais WHERE nofi_nr_Fatura = " & Adodc1.Recordset.Fields("dupl_nr_Fatura") & ")", Estatico
        If EntSaida = "S" Then
            MsgBox "Não é possível alterar esse tipo de documento." + Chr(13) + "Visualize a Glosa através do seu Documento de origem nº " & ResTipoDoc!nofi_nr_documento & ", no Contas a Receber.", vbCritical, "ATENÇÃO"
        Else
            MsgBox "Não é possível alterar esse tipo de documento." + Chr(13) + "Visualize a Devolução através do seu Documento de origem nº " & ResTipoDoc!nofi_nr_documento & ", no Contas a Pagar.", vbCritical, "ATENÇÃO"
        End If
        ResTipoDoc.Close
        Exit Sub
    End If

    'Não permitir efetivação para registros de Estorno Financeiro
    If Left(Adodc1.Recordset.Fields("Numdocumento"), 1) = "E" Then
        MsgBox "Esta opção não é permitida para registro de Estorno Financeiro!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If


    'Não permitir alterar uma transferência de custo:
    If TDBGrid1.Columns("Tipo Doc.") = "TRANSFERÊNCIA DE CUSTO" Then
        MsgBox "Não e possivel alterar uma transferência de custo!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If


    XFormulario = "TelaDuplicatas"
    NFatura = Adodc1.Recordset.Fields("dupl_nr_fatura")
    xChave = Adodc1.Recordset.Bookmark
    ChaveDp = Adodc1.Recordset.Fields("dupl_cd_duplicata")

    'Os registros frutos de um estorno não devem atender a nenhuma solicitação
    If Adodc1.Recordset.Fields("dupl_tx_Estorno") = "S" Then
        MsgBox "Não é permitido alterar um estorno!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If

    SubQOpenRecordset ResFat, "SELECT nofi_tx_controle,nofi_cd_notafiscal,nofi_tx_tipofatura FROM NotasFiscais where nofi_nr_fatura =" & NFatura, "Estatico"
    If Not (ResFat.BOF And ResFat.EOF) Then
        If Not IsNull(ResFat!nofi_tx_controle) Then
            XControleTransf = ResFat!nofi_tx_controle
        End If
        XTipoFatura = ResFat!nofi_tx_tipofatura
        If XTipoFatura <> "3" Then
            Chave = ResFat!nofi_cd_notafiscal
        End If
    End If
    ResFat.Close

    If IsNull(Adodc1.Recordset.Fields("focl_tx_fantasia")) Then
    'Se não existe fornecedor o registro é uma transferencia bancária ou recolhimento
        SubQOpenRecordset ResTipoDoc, "SELECT TipoDocumentos.tido_tx_sigla FROM TipoDocumentos,NotasFiscais WHERE NotasFiscais.nofi_nr_fatura = " & NFatura & " and notasfiscais.tido_cd_tipodoc = TipoDocumentos.tido_cd_tipodoc ", "Estatico"
        If Not ResTipoDoc.EOF Then
            xsigla = FunNulo(ResTipoDoc!tido_tx_sigla)
        Else
            MsgBox "Não existe duplicata com este nº de PP. Atualize a tela!", vbCritical, "ATENÇÃO"
            Exit Sub
        End If
        ResTipoDoc.Close

        If xsigla = "TB" Then
            XFormulario = "TelaDuplicatas"
            FrmTransfBancaria.Show 1
            Set Formulario = TelaDuplicatas
            Exit Sub
        ElseIf (xsigla = "TBCP" Or xsigla = "TBCR") Then
            XFormulario = "TelaNotasFiscais"
            frmTransferenciaComCusto.Show 1
            Set Formulario = TelaNotasFiscais
            Exit Sub

        ElseIf xsigla = "TRCC" Then
            If MsgBox("Não é permitido alterar uma Transferência de Custo. Deseja visualizá-la ?", vbYesNo + vbCritical, "ATENÇÃO") = vbYes Then
                XFormulario = "TelaDuplicatas"
'                FrmTransfCusto.Show 1
                Set Formulario = TelaDuplicatas
                Exit Sub
            Else
                Exit Sub
            End If
        ElseIf xsigla = "DAM" Or xsigla = "DARF" Or xsigla = "INSS" Or xsigla = "OUTR" Then
            MsgBox "Recolhimentos não podem ser alterados. Exclua e efetue novo recolhimento.", vbCritical, "ATENÇÃO"
            Exit Sub
        End If
    End If

    XInserir = "A"
    If Not IsNull(Adodc1.Recordset.Fields("Saida")) Then
        EntSaida = "S"
    ElseIf Not IsNull(Adodc1.Recordset.Fields("Entrada")) Then
        EntSaida = "E"
    End If
    If XTipoFatura = "3" Then
        FrmFatura.Show 1
    Else
        FrmNotasFiscais.Show 1
    End If
    If FrmNotasFiscais.CARREGAGRID = True Then
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
        TDBGrid1.Refresh
        If Not Adodc1.Recordset.EOF Then
            If Adodc1.Recordset.RecordCount > xChave Then
                Adodc1.Recordset.Bookmark = xChave
            End If
        End If
    End If
    
        
        
        
   
End Sub

Private Sub CmdBaixaParcial_Click()
    Dim X As Integer
    
    If Not FunVerifica_Permissao("TelaDuplicatas", "1") Then
        Exit Sub
    End If
    
    'Não permitir efetivação para registros de Estorno Financeiro
    If Left(Adodc1.Recordset.Fields("Numdocumento"), 1) = "E" Then
        MsgBox "Esta opção não é permitida para registro de Estorno Financeiro!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    
    If FunTabelaVazia(Adodc1) Then
        MsgBox "Não pode ser efetivado nenhum pagamento, pois a tabela está vazia", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    If Not IsNull(Adodc1.Recordset.Fields("dupl_dt_pagrec")) And Adodc1.Recordset.Fields("dupl_dt_pagrec") <> "" Then
        MsgBox "O pagamento desta duplicata já foi realizado.", vbInformation, "ATENÇÃO!"
        Exit Sub
    End If
    
    'Os registros frutos de um estorno não devem atender a nenhuma solicitação
    If Adodc1.Recordset.Fields("dupl_tx_Estorno") = "S" Then
        MsgBox "Não é possível efetivar um pagamento para um estorno!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    
    'Não permitir efetivação para transferências de custo:
    If TDBGrid1.Columns("Tipo Doc.") = "TRANSFERÊNCIA DE CUSTO" Then
        MsgBox "Esta opção não é permitida para um registro de transferência de custo!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If

    Chave = Adodc1.Recordset.Fields("dupl_cd_duplicata")
    PanBaixaParcial.Left = (TDBGrid1.Width - PanBaixaParcial.Width) / 2
    PanBaixaParcial.Top = (TDBGrid1.Height - PanBaixaParcial.Height) / 2
    If IsNull(Adodc1.Recordset.Fields("Entrada")) Then
        LblDtPagtoParcial.Caption = "Pagamento:"
        LblTipoResiduo.Caption = "Tipo Pagamento:"
    Else
        LblDtPagtoParcial.Caption = "Recebimento:"
        LblTipoResiduo.Caption = "Tipo Recimento:"
    End If
    PanBaixaParcial.Visible = True
    PanBaixaParcial.Enabled = True
    subDesabilitaBotoes
    TDBGrid1.Enabled = False
    For X = 0 To CboTipoPagRecParcial.ListCount - 1
        If CboTipoPagRecParcial.ItemData(X) = CInt(Adodc1.Recordset.Fields("dupl_tx_tipopagrec")) Then
            CboTipoPagRecParcial.ListIndex = X
        End If
    Next
    
    TxtValorParcial.Text = Format(Adodc1.Recordset.Fields("Valor") + Adodc1.Recordset.Fields("DA"), "standard")
    DtpPagtoParcial.Value = Format$(Now, "DD/MM/YY")
    CboCCorrenteParcial.BoundText = Adodc1.Recordset.Fields("coco_cd_codigo")
    TxtNumDocParcial.Text = ""
    
    DtpVencResiduo.Value = Format$(Adodc1.Recordset.Fields("dupl_dt_vencimento"), "DD/MM/YY")
    CboTipoPagRecResiduo.ListIndex = CboTipoPagRecParcial.ListIndex
    CboCCorrenteResiduo.BoundText = Adodc1.Recordset.Fields("coco_cd_codigo")
    
    If Adodc1.Recordset.Fields("TipoDocDesc") = "FATURA" Then 'Quando se tratar de fatura, o tipo de documento será o das notas fiscais de origem. -10/11/09 - Patrícia
        CboTipoDoc.Text = Adodc1.Recordset.Fields("tido_tx_descricao")
    Else
        CboTipoDoc.Text = Adodc1.Recordset.Fields("TipoDocDesc")
    End If
    DatTipoDoc.Recordset.Bookmark = CboTipoDoc.SelectedItem
    XLI_TIPODOC = CboTipoDoc.BoundText
    
    TxtValorParcial.SetFocus
End Sub

Private Sub CboTipoDoc_Change()
    If CboTipoDoc.Text <> "" Then
        DatTipoDoc.Recordset.Bookmark = CboTipoDoc.SelectedItem
        TxtCodTipoDoc.Text = CboTipoDoc.BoundText
    End If
End Sub


Private Sub CmdCancelar_Click()
    Dim Sql As String
    Dim TotChConsolidado As String
    Dim XTipo As String, xsigla As String
    Dim Xdoc As String
    Dim XConta As String
    Dim XFat As Long
    Dim xcod As Long
    Dim XData As Date
    Dim ResTransf As ADODB.Recordset
    Dim Controle As Double
    Dim XLT_DESCRICAO As String
    
    If Not FunVerifica_Permissao("TelaDuplicatas", "2") Then
        Exit Sub
    End If
    
    'Não permitir efetivação para registros de Estorno Financeiro
    If Left(Adodc1.Recordset.Fields("Numdocumento"), 1) = "E" Then
        MsgBox "Esta opção não é permitida para registro de Estorno Financeiro!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
            
    'Não permitir cancelar um pagamento de glosa ou de devolução
    If TDBGrid1.Columns("Tipo Doc.") = "GLOSA" Or TDBGrid1.Columns("Tipo Doc.") = "NOTA FISCAL DE DEVOLUÇÃO" Then
        MsgBox "Não é possível tornar uma glosa/Devolução pendente!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
        
    'Não permitir cancelar uma transferência de custo:
    If TDBGrid1.Columns("Tipo Doc.") = "TRANSFERÊNCIA DE CUSTO" Then
        MsgBox "Não e possivel cancelar uma transferência de custo!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
        
    
    
    '***************CONTABILIDADE*******************
    'possui o módulo de contabilidade e o tipo de documento é exportável para a contabilidade
    If XGB_PQUALICONT And Adodc1.Recordset.Fields("TipoDocNec") = "0" Then
        If funMesFechado(Adodc1.Recordset.Fields("dupl_dt_PagRec")) <> "A" Then
            If Adodc1.Recordset.Fields("dupl_tx_debcred") = "D" Then
                MsgBox "Não é possível cancelar o Pagamento, pois o mês contábil já foi fechado. " _
                    + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
            Else
                MsgBox "Não é possível cancelar o Recebimento nessa data, pois o mês contábil já foi fechado. " _
                    + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
            End If
            Exit Sub
        End If
    End If
    '***********************************************
    
    subDesabilitaBotoes
    
    'Os registros frutos de um estorno não devem atender a nenhuma solicitação
    If Adodc1.Recordset.Fields("dupl_tx_Estorno") = "S" Then
        MsgBox "Esta opção não é permitida para um estorno!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    
    If Adodc1.Recordset.Fields("dupl_tx_status") <> "R" Then
        MsgBox "Só é possível cancelar entradas/saídas cujos pagamentos já tenham sido realizados.", vbCritical, "ATENÇÃO !!"
        subHabilitaBotoes
        Exit Sub
    End If
    
    If MsgBox("Confirma Cancelamento do Pag./Rec. ?", vbCritical + vbYesNo, "CUIDADO") = vbNo Then
        subHabilitaBotoes
        Exit Sub
    End If
    
    TelaDuplicatas.MousePointer = vbHourglass
            
    xChave = TDBGrid1.Bookmark
    XTipo = Adodc1.Recordset.Fields("dupl_tx_tipopagrec")
    
    'Verifica se é uma transf. Bancária
    If TDBGrid1.Columns("Cliente/Fornecedor") = "" Then
        SubQOpenRecordset ResTransf, "SELECT nofi_tx_controle,tido_tx_sigla FROM ConsFINNotasFiscais where nofi_nr_fatura =" & Adodc1.Recordset.Fields("dupl_nr_fatura"), "Estatico"
        ResTransf.MoveFirst
        xsigla = FunNulo(ResTransf!tido_tx_sigla)
        Controle = FunNuloVal(ResTransf!nofi_tx_controle)
        ResTransf.Close
    End If
        
    Conexao.BeginTrans
    
    If xsigla = "TB" Then
        
        Xdoc = Adodc1.Recordset.Fields("dupl_tx_docpagrec")
        XConta = Adodc1.Recordset.Fields("coco_cd_codigo")
        XData = Adodc1.Recordset.Fields("dupl_dt_pagrec")
        
        SubQOpenRecordset ResTransf, "SELECT dupl_tx_docpagrec,dupl_dt_pagrec," & _
                " dupl_dt_concilia,dupl_tx_status,dupl_cd_Duplicata" & _
                " FROM NotasFiscais left join duplicatas ON duplicatas.dupl_nr_fatura = NotasFiscais.nofi_nr_fatura where nofi_tx_controle ='" & Controle & "' and empr_cd_empresa = " & PCodEmpresa & "", "Dinamico"
                
        If (Not ResTransf.EOF And Not ResTransf.BOF) Then
            
            ResTransf.MoveFirst
            
            Do While Not ResTransf.EOF
                'Atualiza a tabela de duplicatas
                ResTransf!dupl_tx_docpagrec = Null
                ResTransf!dupl_dt_pagrec = Null
                ResTransf!dupl_dt_concilia = Null
                ResTransf!dupl_tx_status = "C"
                                             
                If XTipo = "1" Then
                    TotChConsolidado = Adodc1.Recordset.Fields("valor")
                End If
                
                '***************CONTABILIDADE*******************
                'Exclui a parte da contabilidade
                Conexao.Execute "DELETE FROM Lancamentos " & _
                            " WHERE lanc_cd_OrigemFin= " & ResTransf!dupl_cd_Duplicata & "" & _
                            " AND lanc_tx_Origem = 'FD'" & _
                            " AND empr_cd_empresa = " & PCodEmpresa & ""
                '***********************************************
                ResTransf.MoveNext
            Loop
            
            'Atualiza a tabela de notas fiscais
            Sql = "UPDATE NotasFiscais SET nofi_tx_status='C' WHERE nofi_tx_controle='" & Controle & "'  and empr_cd_empresa = " & PCodEmpresa & ""
            Conexao.Execute Sql
            
            ResTransf.Close
        End If
        
    ElseIf xsigla = "TRCC" Then 'Transferência de Custo
        MsgBox "Esta opção não é permitida para uma Transferência de Custo!", vbCritical, "ATENÇÃO!"
        Exit Sub
        
    ElseIf XTipo = "1" Then '1=CHEQUE;
        Xdoc = Adodc1.Recordset.Fields("dupl_tx_docpagrec")
        XConta = Adodc1.Recordset.Fields("coco_cd_codigo")
        xcod = Adodc1.Recordset.Fields("dupl_cd_duplicata")
        If Not IsNull(Adodc1.Recordset.Fields("dupl_dt_pagrec")) Then XData = Adodc1.Recordset.Fields("dupl_dt_pagrec")
        XFat = Adodc1.Recordset.Fields("dupl_nr_fatura")
        
        'Tornar Pendentes documentos consolidados
        SubQOpenRecordset ResTransf, "SELECT dupl_tx_docpagrec,dupl_dt_pagrec,dupl_tx_status," & _
                " dupl_nr_Fatura,dupl_vl_Valor,dupl_dt_concilia,dupl_cd_Duplicata " & _
                " FROM Duplicatas WHERE dupl_tx_tipopagrec ='" & XTipo & "' and  dupl_tx_docPagRec='" & Xdoc & "' and dupl_tx_status='R' " & _
                " and dupl_dt_PagRec=" & FunNuloData(XData, NomeSgbd) & " AND(dupl_tx_Estorno<>'S' or dupl_tx_Estorno is null)", Dinamico
                
        ResTransf.MoveFirst
        
        Do While Not ResTransf.EOF
            'Atualiza a tabela de duplicatas
            ResTransf!dupl_tx_docpagrec = Null
            ResTransf!dupl_dt_pagrec = Null
            ResTransf!dupl_dt_concilia = Null
            ResTransf!dupl_tx_status = "C"
                                         
            '***************CONTABILIDADE*******************
            'Exclui a parte da contabilidade
            Conexao.Execute "DELETE FROM Lancamentos " & _
                            " WHERE lanc_cd_OrigemFin= " & ResTransf!dupl_cd_Duplicata & "" & _
                            " AND lanc_tx_Origem = 'FD'" & _
                            " AND empr_cd_empresa = " & PCodEmpresa & ""
            '***********************************************
            
            'Atualiza a tabela de notas fiscais
            Sql = "UPDATE NotasFiscais SET nofi_tx_status='C' WHERE nofi_nr_fatura=" & ResTransf!dupl_nr_fatura
            Conexao.Execute Sql
            
            If TotChConsolidado = "" Then
                TotChConsolidado = CStr(ResTransf!dupl_vl_valor)
            Else
                TotChConsolidado = CStr(CDbl(TotChConsolidado) + ResTransf!dupl_vl_valor)
            End If
            
            ResTransf.MoveNext
        Loop
        ResTransf.Close
        
    ElseIf XTipo <> "6" Then
        Xdoc = Adodc1.Recordset.Fields("dupl_tx_docpagrec")
        XConta = Adodc1.Recordset.Fields("coco_cd_codigo")
        xcod = Adodc1.Recordset.Fields("dupl_cd_duplicata")
        If Not IsNull(Adodc1.Recordset.Fields("dupl_dt_pagrec")) Then XData = Adodc1.Recordset.Fields("dupl_dt_pagrec")
        XFat = Adodc1.Recordset.Fields("dupl_nr_fatura")
        
        'Tornar Pendentes documentos consolidados
        'É necessário incluir o estorno, pois quando ocorre um estorno o lançamento origem fica compromissado não sendo possível alcança-lo neste momento
        SubQOpenRecordset ResTransf, _
                "SELECT dupl_tx_docpagrec,dupl_dt_pagrec,dupl_tx_status, dupl_nr_Fatura,dupl_vl_Valor,dupl_dt_concilia,dupl_cd_Duplicata, dupl_tx_Estorno " & _
                " FROM Duplicatas " & _
                " WHERE dupl_tx_tipopagrec ='" & XTipo & _
                    "' AND  dupl_tx_docPagRec='" & Xdoc & _
                    "' AND dupl_tx_status='R' " & _
                    " AND coco_cd_codigo=" & XConta & _
                    " AND dupl_dt_PagRec=" & FunNuloData(XData, NomeSgbd), Dinamico
        
        If (Not ResTransf.EOF And Not ResTransf.BOF) Then
            
            ResTransf.MoveFirst
            
            Do While Not ResTransf.EOF
                If ResTransf!dupl_tx_Estorno <> "S" Or IsNull(ResTransf!dupl_tx_Estorno) Then
                    'Atualiza a tabela de duplicatas
                    ResTransf!dupl_tx_docpagrec = Null
                    ResTransf!dupl_dt_pagrec = Null
                    ResTransf!dupl_dt_concilia = Null
                    ResTransf!dupl_tx_status = "C"
                    
                    'Atualiza a tabela de notas fiscais
                    Sql = "UPDATE NotasFiscais SET nofi_tx_status='C' WHERE nofi_nr_fatura=" & ResTransf!dupl_nr_fatura
                    Conexao.Execute Sql
                    
                    '***************CONTABILIDADE*******************
                    'Exclui a parte da contabilidade
                    Conexao.Execute "DELETE FROM Lancamentos " & _
                                    " WHERE lanc_cd_OrigemFin= " & ResTransf!dupl_cd_Duplicata & "" & _
                                    " AND lanc_tx_Origem = 'FD'" & _
                                    " AND empr_cd_empresa = " & PCodEmpresa & ""
                    '***********************************************
                End If
            
                ResTransf.MoveNext
            Loop
            
            ResTransf.Close
        End If
        
    Else 'DINHEIRO
        'Atualiza a tabela de duplicatas
        Sql = "UPDATE Duplicatas SET dupl_tx_docpagrec=null,dupl_dt_pagrec=null,dupl_tx_status='C' WHERE dupl_cd_duplicata=" & Adodc1.Recordset.Fields("dupl_cd_duplicata")
        Conexao.Execute Sql
        
        'Atualiza a tabela de notas fiscais
        Sql = "UPDATE NotasFiscais SET nofi_tx_status='C' WHERE nofi_nr_fatura=" & Adodc1.Recordset.Fields("dupl_nr_fatura")
        Conexao.Execute Sql
        
        '***************CONTABILIDADE*******************
        'Exclui a parte da contabilidade
        Conexao.Execute "DELETE FROM Lancamentos " & _
                    " WHERE lanc_cd_OrigemFin= " & Adodc1.Recordset.Fields("dupl_cd_duplicata") & "" & _
                    " AND lanc_tx_Origem = 'FD'" & _
                    " AND empr_cd_empresa = " & PCodEmpresa & ""
        '***********************************************
    End If
            
    If XTipo = "1" Then
    'Insere os dados referentes ao cheque cancelado na tabela ChequesCancelados
        Sql = "INSERT INTO ChequesCancelados (chca_tx_numcheque,coco_cd_codigo,chca_dt_emissao,chca_vl_valor)" & _
              " VALUES ('" & Xdoc & "'," & XConta & "," & FunNuloData((XData), NomeSgbd) & "," & FunTrataFloat(TotChConsolidado) & ")"
        Conexao.Execute Sql
    
        Conexao.Execute "DELETE FROM VersoCheques where vech_tx_numcheque= '" & Xdoc & "' and coco_cd_codigo = " & XConta
    End If
      
    
    Conexao.CommitTrans
    
    '*********************REGISTRO DO LOG*************************
    ReDim XGM_MATRIZLOG(3, 2) As Variant
    XGM_MATRIZLOG(0, 0) = "Parcela" 'Campo de identificação da duplicata.
    XGM_MATRIZLOG(1, 0) = "Nº Doc" 'Número de documento.
    XGM_MATRIZLOG(2, 0) = "Valor" 'Campo fatura da duplicata.
    XGM_MATRIZLOG(0, 1) = Adodc1.Recordset.Fields("Parcela") 'Pega a identificação da duplicata.
    XGM_MATRIZLOG(1, 1) = Adodc1.Recordset.Fields("Numdocumento") 'Pega número de documento da nota fiscal que a parcela pertence.
    XGM_MATRIZLOG(2, 1) = Adodc1.Recordset.Fields("Valor") 'Pega o valor da duplicata.
    
    Call subRegistraLog("TelaDuplicatas", "1", funCriaDescricaoLog(XGM_MATRIZLOG, PENDENTE, DUPLICATA))
    
    '*************************************************************
            
    Quantid = Quantid - 1
    MdiPrincipal.BarraStatus.Panels(3).Text = Str(Quantid) + "/" + Str(Quantid)
    
    If ContFil <> 0 Then
       subContaReg Adodc1
    End If

    subHabilitaBotoes
    
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    
    If Not IsNull(TDBGrid1.Bookmark) Then
        TDBGrid1.MoveLast
        If TDBGrid1.Bookmark >= xChave Then
            'Se houve um filtro anteriormente pode ser que este registro não esteja mais no grid
            'Por exemplo se houve um filtro de Realizado e foi cancelado nesse momento o registro não mais estará na tela
            TDBGrid1.Bookmark = xChave
        End If
    End If
    
    TelaDuplicatas.MousePointer = vbDefault
Exit Sub

ResumoDoErro:
    TDBGrid1.SetFocus
    TelaDuplicatas.MousePointer = vbDefault
    Adodc1.Refresh
    TDBGrid1.Bookmark = xChave
    Exit Sub

TrataErro:
    funTrataErros (ComMensagem)
    Resume ResumoDoErro

End Sub

Private Sub cmdCancelarDuplicata_Click()
    
    Dim XLO_CHECK As ADODB.Recordset
    Dim XLT_SQL As String
    'verifica se existe imposto - 12/01/2010 - Patrícia
    XLT_SQL = "SELECT * FROM NOTASFISCAIS INNER JOIN DESCONTOSNF ON (NOTASFISCAIS.NOFI_NR_FATURA = DESCONTOSNF.DEEN_NR_FATURA AND NOFI_NR_FATURA = " & Adodc1.Recordset.Fields("dupl_nr_fatura") & ")"
    SubQOpenRecordset XLO_CHECK, XLT_SQL, Estatico
    If Not (XLO_CHECK.EOF) Then
        MsgBox "Esta opção não é permitida para Notas com impostos relacionados.!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If

    'verifica a mesma permissão do estorno
    If Not FunVerifica_Permissao("TelaDuplicatas", "1") Then
        Exit Sub
    End If
    
    'Não permitir cancelamento de transferências de custo:
    If TDBGrid1.Columns("Tipo Doc.") = "TRANSFERÊNCIA DE CUSTO" Then
        MsgBox "Esta opção não é permitida para um registro de transferência de custo!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    
    'verifica se é compromissado para permitir ou não o cancelamento
    If Adodc1.Recordset.Fields("dupl_tx_status") = "C" Then
        Me.panCancelar.Visible = True
        Me.panCancelar.Left = (TDBGrid1.Width - PanEstorno.Width) / 2
        Me.panCancelar.Top = (TDBGrid1.Height - PanEstorno.Height) / 2
        Me.dtpCancelamento.Value = Now
    'se for nota cancelada perguntar se quer "cancelar" o cancelamento
    ElseIf Adodc1.Recordset.Fields("dupl_tx_status") = "X" Then
        If MsgBox("Parcela já cancelada! Deseja compromissar a duplicata?", vbYesNo + vbCritical, "ATENÇÃO") = vbYes Then
            'chama a sub que realiza o cancelamento do compromisso
            Call subCompromissarCancelamento
            
            
        End If
    Else
        MsgBox "Esta opção só é permitida para duplicatas compromissadas!", vbInformation, "Cancelamento"
        Me.panCancelar.Visible = False
    End If
    
End Sub

Private Sub CmdCanExpBanco_Click()
 
    Dim Sql As String
    Dim TotChConsolidado As String
    Dim XTipo As String, xsigla As String
    Dim Xdoc As String
    Dim XConta As String
    Dim XFat As Integer
    Dim xcod As Integer
    Dim XData As Date
    Dim ResTransf As ADODB.Recordset
    Dim Controle As Integer
    
    'Não permitir efetivação para registros de Estorno Financeiro
    If Left(Adodc1.Recordset.Fields("Numdocumento"), 1) = "E" Then
        MsgBox "Esta opção não é permitida para registro de Estorno Financeiro!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If

    subDesabilitaBotoes
    
    If MsgBox("Confirma Cancelamento da Exportação para o Banco?", vbCritical + vbYesNo, "CUIDADO") = vbNo Then
        subHabilitaBotoes
        Exit Sub
    End If
    TelaDuplicatas.MousePointer = vbHourglass
            
    xChave = TDBGrid1.Bookmark
            
    Conexao.BeginTrans
        
    Conexao.Execute "UPDATE Duplicatas SET dupl_tx_ExpBanco = null WHERE dupl_cd_duplicata = " & Adodc1.Recordset.Fields("dupl_cd_duplicata") & ""
        
    Conexao.CommitTrans
            
    subHabilitaBotoes

    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    If Not IsNull(TDBGrid1.Bookmark) Then
        TDBGrid1.MoveLast
        If TDBGrid1.Bookmark >= xChave Then
            'Se houve um filtro anteriormente pode ser que este registro não esteja mais no grid
            'Por exemplo se houve um filtro de Realizado e foi cancelado nesse momento o registro não mais estará na tela
            TDBGrid1.Bookmark = xChave
        End If
    End If
    TelaDuplicatas.MousePointer = vbDefault
Exit Sub

ResumoDoErro:
    TDBGrid1.SetFocus
    TelaDuplicatas.MousePointer = vbDefault
    Adodc1.Refresh
    TDBGrid1.Bookmark = xChave
    Exit Sub

TrataErro:
    funTrataErros (ComMensagem)
    Resume ResumoDoErro
End Sub
Private Sub cmdCartaTransferencia_Click()
    Dim XLO_TIPODOC As ADODB.Recordset
    Dim XLT_SQL As String
    Dim XLT_HISTORICO As String
    Dim XLO_RSEMPRESA As ADODB.Recordset ' Declaração em 24/09/08 - Patrícia
    
    Modelo = 1 '02/10/08 - Patrícia
    txtEmissor.Enabled = True

    If FunNulo(Adodc1.Recordset.Fields("dupl_tx_status")) <> "R" Then
       MsgBox "O lançamento selecionado não possue data de pagamento!", vbCritical, "ATENÇÃO!"
       Exit Sub
    Else
        DtpDataCartaTransferencia.Value = Adodc1.Recordset.Fields("dupl_dt_PagRec") '29/09/08 - Patrícia
    End If

    
    'Não permitir efetivação para registros de Estorno Financeiro
    If Left(Adodc1.Recordset.Fields("Numdocumento"), 1) = "E" Then
        MsgBox "Esta opção não é permitida para registro de Estorno Financeiro!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    
    If Adodc1.Recordset.Fields("dupl_tx_debCred") = "C" Then

        PanDetalhesCartaTransferencia.Left = (TDBGrid1.Width - PanDetalhesCartaTransferencia.Width) / 2
        PanDetalhesCartaTransferencia.Top = (TDBGrid1.Height - PanDetalhesCartaTransferencia.Height) / 2
                   
        XLT_SQL = "SELECT * FROM Duplicatas WHERE dupl_cd_duplicata = " & Adodc1.Recordset.Fields("dupl_cd_duplicata")
        SubQOpenRecordset XLO_TIPODOC, XLT_SQL, Dinamico
        If IsNull(XLO_TIPODOC!dupl_tx_controle) Then
            PanDetalhesCartaTransferencia.Visible = True
            LbLEmissor.Visible = True
            txtEmissor.Visible = True
            optModelo1.Visible = True
            optModelo2.Visible = True
            optModelo3.Visible = True
        Else
            'If MsgBox("O recibo já foi emitido. Deseja reemitir?", vbYesNo + vbCritical, "ATENÇÃO") = vbYes Then - Documentado em 02/10/2008
                PanDetalhesCartaTransferencia.Visible = True
                LbLEmissor.Visible = True
                txtEmissor.Visible = True
                optModelo1.Visible = True
                optModelo2.Visible = True
                optModelo3.Visible = True
                txtNumeroRecibo.Text = CStr(XLO_TIPODOC!dupl_tx_controle)
                DtpDataCartaTransferencia.Value = CDate(XLO_TIPODOC!dupl_dt_pagrec)
            'Else
            '    Exit Sub
            ' End If
        End If
        XLO_TIPODOC.Close
        
    End If
End Sub

Private Sub CmdConfimarParcial_Click()
    Dim Houve_Erro As Integer  'Indica se Houve Erro
    Dim Sql As String
    Dim ResPadrao As ADODB.Recordset
    Dim StatusDupl As String
    Dim XLF_RESIDUO As Currency     'Valor residual da parcela que está sendo efetivada
    Dim XLT_PARCELA As String       'Código da parcelA ORIGINAL
    Dim XLT_PARCELARESIDUO As String    'Código de Parcela residual
    Dim XLI_CONT As Integer

    Houve_Erro = False
        
    DtpDtPagto.Value = DtpPagtoParcial.Value
    
    'Verificar se o tipo do documento é parcela única
    'Qdo for parcela única, não deve ser permitido baixar parcialmente, porque
    'a baixa parcial gera mais uma duplicata
    DatTipoDoc.Recordset.Bookmark = CboTipoDoc.SelectedItem
    If DatTipoDoc.Recordset.Fields("tido_tx_duplicata") = "1" Then
        MsgBox "Não é possível efetuar Baixa Parcial para um tipo de documento de parcela única. " _
        + Chr(13) + "Altere o tipo de documento para um que permita mais de uma parcela.", vbCritical, "ATENÇÃO"
        CboTipoDoc.SetFocus
        
        Exit Sub
    End If
    
    '***************CONTABILIDADE*******************
    'possui o módulo de contabilidade
    If XGB_PQUALICONT Then
        If funMesFechado(DtpPagtoParcial.Value) <> "A" Then
            If Adodc1.Recordset.Fields("dupl_tx_debcred") = "D" Then
                MsgBox "Não é possível efetivar o Pagamento nessa data, pois o mês contábil já foi fechado. " _
                    + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
            Else
                MsgBox "Não é possível efetivar o Recebimento nessa data, pois o mês contábil já foi fechado. " _
                    + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
            End If
            Exit Sub
        End If
    End If
    '***********************************************
    
    'Checa se os campos do painel foram preenchidos
    xChave = TDBGrid1.Bookmark
    XFT_DOCPAGREC = ""
    
    If CboTipoPagRecParcial.ItemData(CboTipoPagRecParcial.ListIndex) = 1 Then 'cheque
        If FunObrigatorioTXT(TxtNumDocParcial, "O número do documento é obrigatório!") Then
            TxtNumDocParcial.SetFocus
            Exit Sub
        End If
        
        XFT_DOCPAGREC = TxtNumDocParcial.Text
        SubQOpenRecordset ResPadrao, "select dupl_cd_duplicata from duplicatas where dupl_tx_debcred='D' and coco_cd_codigo = " & CboCCorrenteParcial.BoundText & " and dupl_tx_TipoPagRec = '1' and dupl_tx_docpagrec ='" & TxtNumDocParcial.Text & "'", "Estatico"
        If Not ResPadrao.EOF Then
            MsgBox "Este cheque já foi emitido!", vbCritical, "ATENÇÃO!"
            Exit Sub
        End If
        ResPadrao.Close
    ElseIf CboTipoPagRecParcial.ItemData(CboTipoPagRecParcial.ListIndex) = 2 Or CboTipoPagRecParcial.ItemData(CboTipoPagRecParcial.ListIndex) = 3 Or CboTipoPagRecParcial.ItemData(CboTipoPagRecParcial.ListIndex) = 7 Then 'Borderô
        If FunObrigatorioTXT(TxtNumDocParcial, "O número do documento é obrigatório!") Then
            TxtNumDocParcial.SetFocus
            Exit Sub
        End If
        XFT_DOCPAGREC = TxtNumDocParcial.Text
    Else 'crédito, débito em conta ou dinheiro
        If TxtNumDocParcial.Text = "" Then
            XFT_DOCPAGREC = Format(CStr(FunPegaGuardaUltimo("PR")), "A00000")
        Else
            XFT_DOCPAGREC = TxtNumDocParcial.Text
        End If
    End If
    
    XLT_PARCELA = Adodc1.Recordset.Fields("dupl_nr_duplicata")
    '///////////////GRAVAR O RESÍDUO////////////////////
    If Adodc1.Recordset.Fields("Valor") + Adodc1.Recordset.Fields("DA") = CCur(TxtValorParcial.Text) Then
        If MsgBox("O valor da Baixa Parcial é igual ao valor Total da Parcela. Deseja continuar ?", vbInformation + vbYesNo, "ATENÇÃO!") = vbNo Then
            TxtValorParcial.SetFocus
            Exit Sub
        Else
            Conexao.BeginTrans
        End If
    Else
        Conexao.BeginTrans
        
        'GERAR NOVO CÓDIGO DE PARCELAS
        If Adodc1.Recordset.Fields("dupl_nr_duplicata") = "U" Then
            XLT_PARCELA = "A"
            XLT_PARCELARESIDUO = "B"
        Else
            'PEGAR O ÚLTIMO CÓDIGO DE PARCELA PARA INCRMENTAR
            Sql = "SELECT dupl_nr_duplicata FROM duplicatas WHERE dupl_nr_fatura = " & Adodc1.Recordset.Fields("dupl_nr_fatura") & ""
            SubQOpenRecordset ResPadrao, Sql, "Estatico"
            
            If ResPadrao.RecordCount < 199 Then
                If ResPadrao.RecordCount > 0 And ResPadrao.RecordCount < 25 Then
                    XLT_PARCELARESIDUO = Chr(65 + ResPadrao.RecordCount)
                ElseIf ResPadrao.RecordCount >= 25 Then
                    Select Case ResPadrao.RecordCount
                        Case Is <= 34
                            XLT_PARCELARESIDUO = "Z" & Chr(23 + ResPadrao.RecordCount)
                        Case Is <= 44
                            XLT_PARCELARESIDUO = "Z1" & Chr(13 + ResPadrao.RecordCount)
                        Case Is <= 54
                            XLT_PARCELARESIDUO = "Z2" & Chr(3 + ResPadrao.RecordCount)
                        Case Is <= 64
                            XLT_PARCELARESIDUO = "Z3" & Chr(ResPadrao.RecordCount - 7)
                        Case Is <= 74
                            XLT_PARCELARESIDUO = "Z4" & Chr(ResPadrao.RecordCount - 17)
                        Case Is <= 84
                            XLT_PARCELARESIDUO = "Z5" & Chr(ResPadrao.RecordCount - 27)
                        Case Is <= 94
                            XLT_PARCELARESIDUO = "Z6" & Chr(ResPadrao.RecordCount - 37)
                        Case Is <= 104
                            XLT_PARCELARESIDUO = "Z7" & Chr(ResPadrao.RecordCount - 47)
                        Case Is <= 114
                            XLT_PARCELARESIDUO = "Z8" & Chr(ResPadrao.RecordCount - 57)
                        Case Is <= 124
                            XLT_PARCELARESIDUO = "Z9" & Chr(ResPadrao.RecordCount - 67)
                        Case Is <= 134
                            XLT_PARCELARESIDUO = "Z10" & Chr(ResPadrao.RecordCount - 77)
                        Case Is <= 144
                            XLT_PARCELARESIDUO = "Z11" & Chr(ResPadrao.RecordCount - 87)
                        Case Is <= 154
                            XLT_PARCELARESIDUO = "Z12" & Chr(ResPadrao.RecordCount - 97)
                        Case Is <= 164
                            XLT_PARCELARESIDUO = "Z13" & Chr(ResPadrao.RecordCount - 107)
                        Case Is <= 174
                            XLT_PARCELARESIDUO = "Z14" & Chr(ResPadrao.RecordCount - 117)
                        Case Is <= 184
                            XLT_PARCELARESIDUO = "Z15" & Chr(ResPadrao.RecordCount - 127)
                        Case Is <= 194
                            XLT_PARCELARESIDUO = "Z16" & Chr(ResPadrao.RecordCount - 137)
                        Case Is <= 204
                            XLT_PARCELARESIDUO = "Z17" & Chr(ResPadrao.RecordCount - 147)
                    End Select
                End If
            Else
                MsgBox "O número de máximo de parcelas é 75.", vbCritical, "ATENÇÃO"
                Exit Sub
            End If
            ResPadrao.Close
        End If
        
        XLF_RESIDUO = Adodc1.Recordset.Fields("Valor") - CCur(TxtValorParcial.Text) + Adodc1.Recordset.Fields("DA")
        
        Sql = "INSERT INTO Duplicatas(dupl_nr_duplicata," & _
                            "dupl_dt_vencimento," & _
                            "dupl_vl_valor," & _
                            "dupl_tx_status," & _
                            "dupl_tx_tipopagrec," & _
                            "coco_cd_codigo," & _
                            "bafo_cd_codigo," & _
                            "dupl_tx_debcred," & _
                            "dupl_nr_fatura)" & _
               " VALUES ('" & XLT_PARCELARESIDUO & "'," & _
                            "" & FunNuloData(DtpVencResiduo, NomeSgbd) & "," & _
                            "" & FunTrataFloat(XLF_RESIDUO) & "," & _
                            "'C'," & _
                            "'" & CboTipoPagRecResiduo.ItemData(CboTipoPagRecResiduo.ListIndex) & "'," & _
                            "" & CboCCorrenteResiduo.BoundText & "," & _
                            "" & FunNuloVal(Adodc1.Recordset.Fields("bafo_cd_codigo")) & "," & _
                            "'" & Adodc1.Recordset.Fields("dupl_tx_debcred") & "'," & _
                            "" & Str(TDBGrid1.Columns("Fatura")) & ")"
        Conexao.Execute Sql
    End If
    
    Sql = "UPDATE Duplicatas SET dupl_tx_tipopagrec='" & CStr(CboTipoPagRecParcial.ItemData(CboTipoPagRecParcial.ListIndex)) & "', " & _
          "dupl_dt_pagrec =" & FunNuloData(DtpPagtoParcial.Value, NomeSgbd) & ", " & _
          "dupl_tx_docpagrec = '" & FunNulo(XFT_DOCPAGREC) & "', dupl_tx_status='R', " & _
          "coco_cd_codigo='" & CboCCorrenteParcial.BoundText & "'," & _
          "dupl_nr_duplicata = '" & XLT_PARCELA & "'," & _
          "dupl_vl_valor = " & FunTrataFloat(CCur(TxtValorParcial.Text) - Adodc1.Recordset.Fields("DA")) & "" & _
          " WHERE dupl_cd_duplicata=" & Chave
    Conexao.Execute Sql
    subAtualizaNotasFiscais (Adodc1.Recordset.Fields("dupl_nr_fatura"))
    
    'Se alterou o tipo de documento, a nota deve ser atualizada
    If XLI_TIPODOC <> CboTipoDoc.BoundText Then
        Conexao.Execute "UPDATE NotasFiscais SET tido_cd_TipoDoc = " & CboTipoDoc.BoundText & "" & _
                        " WHERE nofi_nR_fatura = " & Adodc1.Recordset.Fields("dupl_nr_fatura") & ""
    End If
    
    Conexao.CommitTrans
    
     
     '****************REGISTRA LOG DA OPERAÇÃO***********************
    
      ReDim XGM_MATRIZLOG(4, 2) As Variant
      
      XGM_MATRIZLOG(0, 0) = "Parcela" 'Identificação da duplicata
      XGM_MATRIZLOG(1, 0) = "Nº Doc" 'Número de documento da nota fiscal da duplicata
      XGM_MATRIZLOG(2, 0) = "Valor" 'Valor da baixa parcial
      XGM_MATRIZLOG(3, 0) = "PP" 'Número de processo de pagamento parcial
      XGM_MATRIZLOG(0, 1) = Adodc1.Recordset.Fields("Parcela")
      XGM_MATRIZLOG(1, 1) = Adodc1.Recordset.Fields("Numdocumento")
      XGM_MATRIZLOG(2, 1) = FunNulo(TxtValorParcial.Text)
      XGM_MATRIZLOG(3, 1) = TxtNumDocParcial.Text
                  
      Call subRegistraLog("TelaDuplicatas", "1", funCriaDescricaoLog(XGM_MATRIZLOG, BAIXA_PARCIAL, DUPLICATA))
      
    '***************************FIM**********************************
    
    
    XNumCheque = FunNulo(TxtNumDocParcial.Text)
    XDtPagamento = DtpPagtoParcial.Value
    
    PanBaixaParcial.Enabled = False
    PanBaixaParcial.Visible = False
    
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    TDBGrid1.Bookmark = xChave
    
    If CboTipoPagRecParcial.ListIndex = 0 And Adodc1.Recordset.Fields("dupl_tx_DebCred") = "D" Then '1:Cheque
        If MsgBox("Deseja imprimir o cheque ?", vbInformation + vbYesNo, "ATENÇÃO!") = vbYes Then
            'Chama a rotina de impressão do cheque
            PanCheque.Enabled = True
            PanCheque.Visible = True
            PanCheque.Left = (TDBGrid1.Width - PanCheque.Width) / 2
            PanCheque.Top = (TDBGrid1.Height - PanCheque.Height) / 2
            TxtNroCheque.Text = XNumCheque
            TxtNroCheque.Enabled = False
            
            SubQOpenRecordset ResPadrao, "SELECT Empr_tx_cidade FROM Empresas", "Estatico"
            TxtCidade.Text = ResPadrao!empr_tx_cidade
            ResPadrao.Close
            
            If Not IsNull(Adodc1.Recordset.Fields("focl_tx_razaosocial")) Then CboNominal.Text = Adodc1.Recordset.Fields("focl_tx_razaosocial")
            TxtVerso.Text = TDBGrid1.Columns("Cliente/Fornecedor") + " - " + xnota + " - " + Xhistorico
            TxtCidade.SetFocus
            Exit Sub
        Else
            TDBGrid1.Enabled = True
            subHabilitaBotoes
            
            subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
            TDBGrid1.Bookmark = xChave
            Exit Sub
        End If
    End If
   
    TDBGrid1.Enabled = True
    subHabilitaBotoes
    
    TDBGrid1.SetFocus
    

Exit Sub
    
RotuloErro:

    subDesabilitaBotoes
    Houve_Erro = True
    Select Case funTrataErros(ComMensagem)

      Case 3000
         MsgBox Campos_Obrig + " são obrigatórios", 48, "ERRO"
          Resume Next

      Case ErrChaveDuplicada
            Resume Next

      Case ErrRegistroExcluido
            Adodc1.Refresh
            Resume Next

      Case 3200
         MsgBox "Código não pode ser alterado pois existem outras tabelas vinculadas a ele.", 64, "ERRO"
          Resume Next

    End Select

End Sub

Private Sub CmdConfirmaduplicata_Click()
    Dim XLT_SQL As String

    xChave = TDBGrid1.Bookmark
    
    Conexao.BeginTrans
    
    XLT_SQL = "UPDATE Duplicatas SET dupl_tx_tipopagrec='" & CStr(CboTipoPagto2.ItemData(CboTipoPagto2.ListIndex)) & "', " & _
          "dupl_dt_vencimento =" & FunNuloData(DtpVencimento.Value, NomeSgbd) & ", " & _
          "coco_cd_codigo='" & CboCCorrente3.BoundText & "' " & _
          "WHERE dupl_cd_duplicata=" & Adodc1.Recordset.Fields("dupl_cd_duplicata")
    Conexao.Execute XLT_SQL
    
    Conexao.CommitTrans
    
    '****************REGISTRA LOG DA OPERAÇÃO***********************
      ReDim XGM_MATRIZLOG(6, 3) As Variant
      XGM_MATRIZLOG(0, 0) = "Parcela" 'Campo de identificação da duplicata.
      XGM_MATRIZLOG(1, 0) = "Nº Doc" 'Campo Número de documento.
      XGM_MATRIZLOG(2, 0) = "Valor" 'Campo fatura da duplicata.
      XGM_MATRIZLOG(3, 0) = "Data Vencimento" 'Campo data de vencimento da duplicata
      XGM_MATRIZLOG(4, 0) = "Tipo de Pagamento" 'Campo forma de pagamento da duplicata (cheque, dinheiro...)
      XGM_MATRIZLOG(5, 0) = "Conta Corrente" 'Campo conta corrente para o pagamento da duplicata
      XGM_MATRIZLOG(0, 1) = Adodc1.Recordset.Fields("Parcela") 'Pega a identificação da duplicata.
      XGM_MATRIZLOG(1, 1) = Adodc1.Recordset.Fields("Numdocumento") 'Pega número de documento da nota fiscal que a parcela pertence.
      XGM_MATRIZLOG(2, 1) = Adodc1.Recordset.Fields("Valor") 'Pega o valor da duplicata.
      XGM_MATRIZLOG(3, 1) = Adodc1.Recordset.Fields("dupl_dt_vencimento") 'Pega a data de vencimento atual da duplicata
      XGM_MATRIZLOG(4, 1) = Adodc1.Recordset.Fields("TipoPagto") 'Pega o tipo de pagamento atual da duplicata
      XGM_MATRIZLOG(5, 1) = Adodc1.Recordset.Fields("BAEmpresa") 'Pega a conta corrente atual p/ pagamento da duplicata
      XGM_MATRIZLOG(0, 2) = Adodc1.Recordset.Fields("Parcela")
      XGM_MATRIZLOG(1, 2) = Adodc1.Recordset.Fields("Numdocumento")
      XGM_MATRIZLOG(2, 2) = Adodc1.Recordset.Fields("Valor")
      XGM_MATRIZLOG(3, 2) = DtpVencimento.Value 'Pega a nova data de vencimento da duplicata
      XGM_MATRIZLOG(4, 2) = CboTipoPagto2.List(CboTipoPagto2.ListIndex) 'Pega o novo tipo de pagamento da duplicata
      XGM_MATRIZLOG(5, 2) = CboCCorrente3.Text 'Pega a nova conta corrente da duplicata
                      
      Call subRegistraLog("TelaDuplicatas", "1", funCriaDescricaoLog(XGM_MATRIZLOG, ALTERAR, DUPLICATA))
      
    ' *************************************************************

    
    PanDuplicata.Enabled = False
    PanDuplicata.Visible = False
    TDBGrid1.Enabled = True
    subHabilitaBotoes
    
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    TDBGrid1.Bookmark = xChave
    TDBGrid1.SetFocus
End Sub

Private Sub CmdConfirmaEstorno_Click()
    Dim XLO_ESTORNO As Object
    Dim XLL_CODESTORNO As Long
    Dim XLB_CONTABILIZOU As Boolean
    Dim XLO_DESCACRES As Object
    Dim XGT_SQL As String
    Dim XLT_DUPLICATA As String
    Dim Sql As String
    
    On Error GoTo TrataErro
    
    Conexao.BeginTrans
    
    Sql = "INSERT INTO Duplicatas(dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,coco_cd_codigo,bafo_cd_codigo,dupl_tx_debcred,dupl_nr_fatura,dupl_dt_pagrec,dupl_tx_Estorno)" & _
          " VALUES ('" & Right(Adodc1.Recordset.Fields("parcela"), 1) & "'," & FunNuloData(TDBGrid1.Columns("Vencimento"), NomeSgbd) & ", " & (Str(-Adodc1.Recordset.Fields("Valor"))) & ",'" & Adodc1.Recordset.Fields("dupl_tx_Status") & "'," & _
          "'" & Trim(Str(Adodc1.Recordset.Fields("dupl_tx_tipopagrec"))) & "','" & FunNulo(TDBGrid1.Columns("Nº Doc. Pagto./Rec.")) & "'," & Adodc1.Recordset.Fields("coco_cd_codigo") & "," & FunNuloVal(Adodc1.Recordset.Fields("bafo_cd_codigo")) & "," & _
          "'" & Adodc1.Recordset.Fields("dupl_tx_debcred") & "'," & Str(TDBGrid1.Columns("Fatura")) & "," & FunNuloData(DtpDtEstorno.Value, NomeSgbd) & ", 'S')"
    Conexao.Execute Sql
    'Verificar data de conciliação(SIM ou NÃO?)
    
 
    'Para estorno com descontos/acréscimos Em: 24/03/06
    If TDBGrid1.Columns("Desconto/Acréscimo").Text <> 0# Then
          'Pega o código da duplicata que acabou de ser inserida
        XGT_SQL = "SELECT dupl_cd_duplicata FROM Duplicatas " & _
        " WHERE dupl_tx_docpagrec = " & "'" & FunNulo(TDBGrid1.Columns("Nº Doc. Pagto./Rec.")) & "'" & _
        " AND coco_cd_codigo = " & Adodc1.Recordset.Fields("coco_cd_codigo") & " AND dupl_vl_valor = " & (Str(-Adodc1.Recordset.Fields("Valor")))
        SubQOpenRecordset XLO_DESCACRES, XGT_SQL, Estatico

        XLT_DUPLICATA = XLO_DESCACRES!dupl_cd_Duplicata

        XLO_DESCACRES.Close
    
        XGT_SQL = ""

        'Pega os descontos da duplicata origem e insere no dupl do estorno
        XGT_SQL = "SELECT deac_cd_descacresc, dude_vl_descacresc FROM DuplDescAcresc " & _
        " WHERE dupl_cd_duplicata = " & Adodc1.Recordset.Fields("dupl_cd_duplicata")
        SubQOpenRecordset XLO_DESCACRES, XGT_SQL, Estatico

       While Not XLO_DESCACRES.EOF
            Sql = "INSERT INTO DuplDescAcresc (dupl_cd_duplicata, deac_cd_descacresc, dude_vl_descacresc)" & _
                    " VALUES ('" & XLT_DUPLICATA & "', " & "'" & XLO_DESCACRES!deac_cd_descacresc & "'" & ", " & _
                    "'" & Str(-XLO_DESCACRES!dude_vl_descacresc) & "'" & ")"
            Conexao.Execute Sql
            XLO_DESCACRES.MoveNext
       Wend
       XLO_DESCACRES.Close
    End If
    
    If TDBGrid1.Columns("Tipo Pagto./Rec.").Text <> "Cheque" Then
        'Insere a data de conciliação
        
        SubQOpenRecordset XLO_ESTORNO, "SELECT * FROM Duplicatas WHERE dupl_cd_duplicata= -1", Dinamico
        XLO_ESTORNO.AddNew

        XLO_ESTORNO!dupl_nr_duplicata = Right(Adodc1.Recordset.Fields("parcela"), 1)
        XLO_ESTORNO!dupl_dt_vencimento = TDBGrid1.Columns("Vencimento")
        XLO_ESTORNO!dupl_vl_valor = Adodc1.Recordset.Fields("Valor")
        XLO_ESTORNO!dupl_tx_status = Adodc1.Recordset.Fields("dupl_tx_Status")
        XLO_ESTORNO!dupl_tx_tipopagrec = Trim(Str(Adodc1.Recordset.Fields("dupl_tx_tipopagrec")))
        XLO_ESTORNO!dupl_tx_docpagrec = FunNulo(TDBGrid1.Columns("Nº Doc. Pagto./Rec."))
        XLO_ESTORNO!coco_cd_codigo = Adodc1.Recordset.Fields("coco_cd_codigo")
        XLO_ESTORNO!bafo_cd_codigo = FunNuloVal(Adodc1.Recordset.Fields("bafo_cd_codigo"))
        XLO_ESTORNO!dupl_tx_debcred = Adodc1.Recordset.Fields("dupl_tx_debcred")
        XLO_ESTORNO!dupl_nr_fatura = Str(TDBGrid1.Columns("Fatura"))
        XLO_ESTORNO!dupl_dt_pagrec = TDBGrid1.Columns("Pagto./Rec.")
        XLO_ESTORNO!dupl_tx_Estorno = "S"
        XLO_ESTORNO!dupl_dt_concilia = Adodc1.Recordset.Fields("dupl_dt_concilia")
        XLO_ESTORNO!dupl_dt_ProcPagamento = Adodc1.Recordset.Fields("dupl_dt_ProcPagamento")
        
        XLO_ESTORNO.Update
        
        XLL_CODESTORNO = XLO_ESTORNO!dupl_cd_Duplicata
        
         'Inserindo os desc/acres para a nova duplicata de estorno
        If TDBGrid1.Columns("Desconto/Acréscimo").Text <> 0# Then
                 'Pega os descontos da duplicata origem e insere no dupl do estorno
            XGT_SQL = "SELECT deac_cd_descacresc, dude_vl_descacresc FROM DuplDescAcresc " & _
                     " WHERE dupl_cd_duplicata = " & Adodc1.Recordset.Fields("dupl_cd_duplicata")
            SubQOpenRecordset XLO_DESCACRES, XGT_SQL, Estatico
        
            While Not XLO_DESCACRES.EOF
                Sql = "INSERT INTO DuplDescAcresc (dupl_cd_duplicata, deac_cd_descacresc, dude_vl_descacresc)" & _
                        " VALUES ('" & XLL_CODESTORNO & "', " & "'" & XLO_DESCACRES!deac_cd_descacresc & "'" & ", " & _
                        "'" & Str(XLO_DESCACRES!dude_vl_descacresc) & "'" & ")"
                Conexao.Execute Sql
                XLO_DESCACRES.MoveNext
            Wend
            XLO_DESCACRES.Close
        End If
                
    XLO_ESTORNO.Close
        
        '***************CONTABILIDADE*******************
        'XLB_CONTABILIZOU = FunContabilizarEstorno(XLL_CODESTORNO, Adodc1.Recordset.Fields("dupl_tx_tipopagrec"), Adodc1.Recordset.Fields("dupl_tx_debcred"), Adodc1.Recordset.Fields("dupl_dt_pagrec"), Adodc1.Recordset.Fields("dupl_nr_fatura"))
'        XLB_CONTABILIZOU = FunContabilizarEstorno(XLL_CODESTORNO, Adodc1.Recordset.Fields("dupl_tx_tipopagrec"), Adodc1.Recordset.Fields("dupl_tx_debcred"), DtpDtEstorno.Value, Adodc1.Recordset.Fields("dupl_nr_fatura"))
        
        'Alterar o código da duplicata para o mesmo código da parcela estorno (+)
        If XLB_CONTABILIZOU = True Then
            
            SubQOpenRecordset XLO_ESTORNO, "SELECT * FROM Lancamentos WHERE lanc_cd_OrigemFin = " & Adodc1.Recordset.Fields("dupl_cd_duplicata") & "", Estatico
            If Not XLO_ESTORNO.EOF Then
                If funMesFechado(Adodc1.Recordset.Fields("dupl_dt_pagrec")) <> "A" Then
                'Abrir Mês para atualizar fatura
                    Conexao.Execute "UPDATE Fechamento SET fech_tx_status = 'A' where fech_dt_mes = '" & Month(Adodc1.Recordset.Fields("dupl_dt_pagrec")) & "/01/" & Year(Adodc1.Recordset.Fields("dupl_dt_pagrec")) & "' AND empr_cd_empresa = " & PCodEmpresa
                    Conexao.Execute "UPDATE Lancamentos SET lanc_cd_OrigemFin = " & XLL_CODESTORNO & " WHERE lanc_cd_OrigemFin = " & Adodc1.Recordset.Fields("dupl_cd_duplicata") & ""
                    Conexao.Execute "UPDATE Fechamento SET fech_tx_status = 'F' where fech_dt_mes = '" & Month(Adodc1.Recordset.Fields("dupl_dt_pagrec")) & "/01/" & Year(Adodc1.Recordset.Fields("dupl_dt_pagrec")) & "' AND empr_cd_empresa = " & PCodEmpresa
                Else
                    Conexao.Execute "UPDATE Lancamentos SET lanc_cd_OrigemFin = " & XLL_CODESTORNO & " WHERE lanc_cd_OrigemFin = " & Adodc1.Recordset.Fields("dupl_cd_duplicata") & ""
                End If
            End If
            XLO_ESTORNO.Close
        End If
        '***********************************************
        'Atualiza a tabela de duplicatas
        Sql = "UPDATE Duplicatas SET dupl_tx_docpagrec=null,dupl_dt_pagrec=null,dupl_tx_status='C', dupl_dt_concilia=null, dupl_dt_ProcPagamento=null WHERE dupl_cd_duplicata=" & Adodc1.Recordset.Fields("dupl_cd_duplicata")
        Conexao.Execute Sql
        
        'Atualiza a tabela de notas fiscais
        Sql = "UPDATE NotasFiscais SET nofi_tx_status='C' WHERE nofi_nr_fatura=" & Adodc1.Recordset.Fields("dupl_nr_fatura")
        Conexao.Execute Sql
    Else
        SubQOpenRecordset XLO_ESTORNO, "SELECT * FROM Duplicatas WHERE dupl_cd_duplicata= -1", Dinamico
        XLO_ESTORNO.AddNew

        XLO_ESTORNO!dupl_nr_duplicata = Right(Adodc1.Recordset.Fields("parcela"), 1)
        XLO_ESTORNO!dupl_dt_vencimento = TDBGrid1.Columns("Vencimento")
        XLO_ESTORNO!dupl_vl_valor = Adodc1.Recordset.Fields("Valor")
        XLO_ESTORNO!dupl_tx_status = Adodc1.Recordset.Fields("dupl_tx_Status")
        XLO_ESTORNO!dupl_tx_tipopagrec = Trim(Str(Adodc1.Recordset.Fields("dupl_tx_tipopagrec")))
        XLO_ESTORNO!dupl_tx_docpagrec = FunNulo(TDBGrid1.Columns("Nº Doc. Pagto./Rec."))
        XLO_ESTORNO!coco_cd_codigo = Adodc1.Recordset.Fields("coco_cd_codigo")
        XLO_ESTORNO!bafo_cd_codigo = FunNuloVal(Adodc1.Recordset.Fields("bafo_cd_codigo"))
        XLO_ESTORNO!dupl_tx_debcred = Adodc1.Recordset.Fields("dupl_tx_debcred")
        XLO_ESTORNO!dupl_nr_fatura = Str(TDBGrid1.Columns("Fatura"))
        XLO_ESTORNO!dupl_dt_pagrec = TDBGrid1.Columns("Pagto./Rec.")
        XLO_ESTORNO!dupl_tx_Estorno = "S"
        
        XLO_ESTORNO.Update
        
        XLL_CODESTORNO = XLO_ESTORNO!dupl_cd_Duplicata
        
        XLO_ESTORNO.Close
        
        
        If TDBGrid1.Columns("Desconto/Acréscimo").Text <> 0# Then
                 'Pega os descontos da duplicata origem e insere no dupl do estorno
            XGT_SQL = "SELECT deac_cd_descacresc, dude_vl_descacresc FROM DuplDescAcresc " & _
                     " WHERE dupl_cd_duplicata = " & Adodc1.Recordset.Fields("dupl_cd_duplicata")
            SubQOpenRecordset XLO_DESCACRES, XGT_SQL, Estatico
            
            While Not XLO_DESCACRES.EOF
                Sql = "INSERT INTO DuplDescAcresc (dupl_cd_duplicata, deac_cd_descacresc, dude_vl_descacresc)" & _
                        " VALUES ('" & XLL_CODESTORNO & "', " & "'" & XLO_DESCACRES!deac_cd_descacresc & "'" & ", " & _
                        "'" & Str(XLO_DESCACRES!dude_vl_descacresc) & "'" & ")"
                Conexao.Execute Sql
                XLO_DESCACRES.MoveNext
            Wend
            XLO_DESCACRES.Close
        End If
        
        
        '***************CONTABILIDADE*******************
'        XLB_CONTABILIZOU = FunContabilizarEstorno(XLL_CODESTORNO, Adodc1.Recordset.Fields("dupl_tx_tipopagrec"), Adodc1.Recordset.Fields("dupl_tx_debcred"), Adodc1.Recordset.Fields("dupl_dt_pagrec"), Adodc1.Recordset.Fields("dupl_nr_fatura"))
        
        'Alterar o código da duplicata para o mesmo código da parcela estorno (+)
        If XLB_CONTABILIZOU = True Then
            Conexao.Execute "UPDATE Lancamentos SET lanc_cd_OrigemFin = " & XLL_CODESTORNO & " WHERE lanc_cd_OrigemFin = " & Adodc1.Recordset.Fields("dupl_cd_duplicata") & ""
        End If
        
        'Lançar o 3º registro na contabilidade, prevendo a reapresentação do cheque
'        SubContabilizarParcela Adodc1.Recordset.Fields("dupl_cd_Duplicata"), Adodc1.Recordset.Fields("dupl_dt_pagrec"), False, False
        '***********************************************
    End If
    
    Conexao.CommitTrans
        
    '****************REGISTRA LOG DA OPERAÇÃO***********************
      ReDim XGM_MATRIZLOG(3, 2) As Variant
      XGM_MATRIZLOG(0, 0) = "Parcela" 'Campo de identificação da duplicata.
      XGM_MATRIZLOG(1, 0) = "Nº Doc" 'Número de documento.
      XGM_MATRIZLOG(2, 0) = "Valor" 'Campo fatura da duplicata.
      XGM_MATRIZLOG(0, 1) = Adodc1.Recordset.Fields("Parcela") 'Pega a identificação da duplicata.
      XGM_MATRIZLOG(1, 1) = Adodc1.Recordset.Fields("Numdocumento") 'Pega número de documento da nota fiscal que a parcela pertence.
      XGM_MATRIZLOG(2, 1) = Adodc1.Recordset.Fields("Valor") 'Pega o valor da duplicata.
                
      Call subRegistraLog("TelaDuplicatas", "1", funCriaDescricaoLog(XGM_MATRIZLOG, ESTORNO, DUPLICATA))
      
    ' *************************************************************
      
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    
    If Not IsNull(TDBGrid1.Bookmark) Then
        If Sql <> "" Then 'ou seja, não é um cancelamento e sim geração de estorno
            If Posicao <= Adodc1.Recordset.RecordCount Then
                TDBGrid1.Bookmark = Posicao
            End If
        ElseIf Posicao > 2 Then
            'se tiver sido cancelado o bookmark será apontado para dois registros (positivo e negativo) acima, já que estes serão excluidos da tela
            TDBGrid1.Bookmark = (Posicao - 2)
        End If
    End If
    
    PanEstorno.Enabled = False
    PanEstorno.Visible = False
    
    Exit Sub

TrataErro:
    Conexao.RollbackTrans
    funTrataErros (ComMensagem)
   
    
End Sub

Private Sub CmdConfirmar_Click()
    Dim Houve_Erro As Integer  'Indica se Houve Erro
    Dim Sql As String
    Dim ResPadrao As ADODB.Recordset
    Dim StatusDupl As String

    Houve_Erro = False
    
    '***************CONTABILIDADE*******************
    'possui o módulo de contabilidade
    If XGB_PQUALICONT Then
        If funMesFechado(DtpDtPagto.Value) <> "A" Then
            If Adodc1.Recordset.Fields("dupl_tx_debcred") = "D" Then
                MsgBox "Não é possível efetivar o Pagamento nessa data, pois o mês contábil já foi fechado. " _
                    + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
            Else
                MsgBox "Não é possível efetivar o Recebimento nessa data, pois o mês contábil já foi fechado. " _
                    + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
            End If
            Exit Sub
        End If
    End If
    '***********************************************
    
    'Checa se os campos do painel foram preenchidos
    xChave = TDBGrid1.Bookmark
    XFT_DOCPAGREC = ""
    'CHEQUE
    If CboTipoPag.ItemData(CboTipoPag.ListIndex) = 1 Then
        If FunObrigatorioTXT(TxtNumDoc, "O número do documento é obrigatório!") Then
            TxtNumDoc.SetFocus
            Exit Sub
        End If
        
        XFT_DOCPAGREC = TxtNumDoc.Text
        'Verifica se o cheque já foi emitido
        SubQOpenRecordset ResPadrao, "select dupl_cd_duplicata from duplicatas where dupl_tx_debcred='D' and coco_cd_codigo = " & CboCCorrente.BoundText & " and dupl_tx_TipoPagRec = '1' and dupl_tx_docpagrec ='" & TxtNumDoc.Text & "'", "Estatico"
        If Not ResPadrao.EOF Then
            MsgBox "Este cheque já foi emitido!", vbCritical, "ATENÇÃO!"
            Exit Sub
        End If
        ResPadrao.Close
    'BORDERÔ
    ElseIf CboTipoPag.ItemData(CboTipoPag.ListIndex) = 2 Or CboTipoPag.ItemData(CboTipoPag.ListIndex) = 3 Or CboTipoPag.ItemData(CboTipoPag.ListIndex) = 7 Then
        XFT_DOCPAGREC = CStr(FunPegaGuardaUltimo("B"))
        TxtNumDoc.Text = XFT_DOCPAGREC
    'CRÉDITO, DÉBITO EM CONTA OU DINHEIRO
    Else
        If TxtNumDoc.Text = "" Then
            XFT_DOCPAGREC = Format(CStr(FunPegaGuardaUltimo("PR")), "A00000")
        Else
            XFT_DOCPAGREC = TxtNumDoc.Text
            'Verifica se o doc já foi registrado
            If CboTipoPag.ItemData(CboTipoPag.ListIndex) = 5 Then
                'Verifica se o nº do doc. pagto. já foi emitido
                SubQOpenRecordset ResPadrao, "SELECT dupl_cd_duplicata FROM duplicatas WHERE dupl_tx_debcred='C' AND coco_cd_codigo = " & CboCCorrente.BoundText & " AND dupl_tx_TipoPagRec = '5' AND dupl_tx_docpagrec ='" & TxtNumDoc.Text & "'  AND dupl_dt_pagrec = " & FunNuloData(DtpDtPagto.Value, NomeSgbd) & "", "Estatico"
                If Not ResPadrao.EOF Then
                    MsgBox "Este nº doc. de crédito já foi registrado!", vbCritical, "ATENÇÃO!"
                    Exit Sub
                End If
                ResPadrao.Close
            ElseIf CboTipoPag.ItemData(CboTipoPag.ListIndex) = 4 Then
                'Verifica se o nº do doc. pagto. já foi emitido
                SubQOpenRecordset ResPadrao, "SELECT dupl_cd_duplicata FROM duplicatas WHERE dupl_tx_debcred='D' AND coco_cd_codigo = " & CboCCorrente.BoundText & " AND dupl_tx_TipoPagRec = '4' AND dupl_tx_docpagrec ='" & TxtNumDoc.Text & "' AND dupl_dt_pagrec = " & FunNuloData(DtpDtPagto.Value, NomeSgbd) & "", "Estatico"
                If Not ResPadrao.EOF Then
                    MsgBox "Este nº doc. de débito já foi registrado!", vbCritical, "ATENÇÃO!"
                    Exit Sub
                End If
                ResPadrao.Close
            End If
        End If
    End If
    
    Conexao.BeginTrans
    
    Sql = "UPDATE Duplicatas SET dupl_tx_tipopagrec='" & CStr(CboTipoPag.ItemData(CboTipoPag.ListIndex)) & "', " & _
          "dupl_dt_pagrec =" & FunNuloData(DtpDtPagto.Value, NomeSgbd) & ", " & _
          "dupl_tx_docpagrec = '" & FunNulo(XFT_DOCPAGREC) & "', dupl_tx_status='R', " & _
          "coco_cd_codigo= '" & CboCCorrente.BoundText & "', " & _
          "fopa_cd_codigo= " & FunNuloValor(CboFormaPagamento) & " " & _
          "WHERE dupl_cd_duplicata=" & Chave
    Conexao.Execute Sql
    subAtualizaNotasFiscais (Adodc1.Recordset.Fields("dupl_nr_fatura"))
    
    '******************REGISTRA LOG DA OPERAÇÃO**************************
      ReDim XGM_MATRIZLOG(4, 2) As Variant
      XGM_MATRIZLOG(0, 0) = "Parcela" 'Campo de identificação da duplicata.
      XGM_MATRIZLOG(1, 0) = "Nº Doc" 'Número de documento.
      XGM_MATRIZLOG(2, 0) = "Valor" 'Campo fatura da duplicata.
      XGM_MATRIZLOG(3, 0) = "Doc. Pagto/Rec" 'Número do processo de pagamento.
      XGM_MATRIZLOG(0, 1) = Adodc1.Recordset.Fields("Parcela") 'Pega a identificação da duplicata.
      XGM_MATRIZLOG(1, 1) = Adodc1.Recordset.Fields("Numdocumento") 'Pega número de documento da nota fiscal que a parcela pertence.
      XGM_MATRIZLOG(2, 1) = Adodc1.Recordset.Fields("Valor") 'Pega o valor da duplicata.
      XGM_MATRIZLOG(3, 1) = FunNulo(XFT_DOCPAGREC) 'Pega número de documento de pagamento da duplicata.
      
      Call subRegistraLog("TelaDuplicatas", "1", funCriaDescricaoLog(XGM_MATRIZLOG, EFETIVAR, DUPLICATA))
      
    '*******************************************************************
    
    Conexao.CommitTrans
    
    XNumCheque = FunNulo(TxtNumDoc.Text)
    XDtPagamento = DtpDtPagto.Value
    
    PanPagamento.Enabled = False
    PanPagamento.Visible = False
    
    If CboTipoPag.ListIndex = 0 And Adodc1.Recordset.Fields("dupl_tx_DebCred") = "D" Then '1:Cheque
        If MsgBox("Deseja imprimir o cheque ?", vbInformation + vbYesNo, "ATENÇÃO!") = vbYes Then
            'Chama a rotina de impressão do cheque
            PanCheque.Enabled = True
            PanCheque.Visible = True
            PanCheque.Left = (TDBGrid1.Width - PanCheque.Width) / 2
            PanCheque.Top = (TDBGrid1.Height - PanCheque.Height) / 2
            TxtNroCheque.Text = XNumCheque
            TxtNroCheque.Enabled = False
            
            SubQOpenRecordset ResPadrao, "SELECT Empr_tx_cidade FROM Empresas", "Estatico"
            TxtCidade.Text = ResPadrao!empr_tx_cidade
            ResPadrao.Close
            
            If Not IsNull(Adodc1.Recordset.Fields("focl_tx_razaosocial")) Then CboNominal.Text = Adodc1.Recordset.Fields("focl_tx_razaosocial")
            TxtVerso.Text = TDBGrid1.Columns("Cliente/Fornecedor") + " - " + xnota + " - " + Xhistorico
            TxtCidade.SetFocus
            Exit Sub
        Else
            TDBGrid1.Enabled = True
            subHabilitaBotoes
            subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
            TDBGrid1.Bookmark = xChave
            Exit Sub
        End If
    End If
    
    TDBGrid1.Enabled = True
    subHabilitaBotoes
    
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    TDBGrid1.Bookmark = xChave
    TDBGrid1.SetFocus
Exit Sub
    
RotuloErro:
  subDesabilitaBotoes
  Houve_Erro = True
  Select Case funTrataErros(ComMensagem)
  Case 3000
     MsgBox Campos_Obrig + " são obrigatórios", 48, "ERRO"
      Resume Next
  Case ErrChaveDuplicada
        Resume Next
  Case ErrRegistroExcluido
        Adodc1.Refresh
        Resume Next
  Case 3200
     MsgBox "Código não pode ser alterado pois existem outras tabelas vinculadas a ele.", 64, "ERRO"
      Resume Next
  End Select
End Sub

Sub subAtualizaNotasFiscais(FaturaDupl As Long)
  Dim StatusDupl As String
  Dim ResNF As ADODB.Recordset
  Dim xsigla As String, XControle As String
  
  StatusDupl = ""
  'Verifica Se todas as duplicatas de uma nota já forma pagas. Se sim atualiza a nota
  SubQOpenRecordset ResNF, "SELECT dupl_tx_status FROM Duplicatas where dupl_nr_fatura =" & FaturaDupl, "Estatico"
  ResNF.MoveFirst
  While Not ResNF.EOF
     If ResNF!dupl_tx_status <> "R" Then
       StatusDupl = ResNF!dupl_tx_status
     End If
     ResNF.MoveNext
  Wend
  ResNF.Close
  If StatusDupl = "" Then
    StatusDupl = "R"
  End If
  
  SubQOpenRecordset ResNF, _
    "SELECT nofi_tx_controle, " & _
    "  nofi_tx_status, " & _
    "  tido_tx_sigla, " & _
    "  nofi_tx_tipofatura, " & _
    "  nofi_tx_historico, " & _
    "  nofi_nr_documento, " & _
    "  nofi_nr_fatura " & _
    "FROM ConsFINNotasFiscais " & _
    "WHERE nofi_nr_fatura =" & FaturaDupl, _
    "Estatico"
  ResNF.MoveFirst
  xsigla = FunNulo(ResNF!tido_tx_sigla)
  XControle = FunNulo(ResNF!nofi_tx_controle)
  If ResNF!nofi_tx_tipofatura = "3" Then
     Xhistorico = "Fatura"
     xnota = CStr(ResNF!nofi_nr_fatura)
  ElseIf ResNF!nofi_tx_tipofatura = "1" Or ResNF!nofi_tx_tipofatura = "2" Then
     Xhistorico = FunNulo(ResNF!nofi_tx_historico)
     xnota = ResNF!nofi_nr_documento
  End If
  If StatusDupl = "R" Then Conexao.Execute _
    "UPDATE NotasFiscais SET " & _
    "  nofi_tx_status = '" & StatusDupl & "' " & _
    "WHERE nofi_nr_fatura =" & FaturaDupl
  ResNF.Close
  
  'Atualizar as partes da transferência bancária
  'Se não existe fornecedor o registro é uma transferencia bancária ou um recolhimento
  If TDBGrid1.Columns("Cliente/Fornecedor").Value = "" Then
    If xsigla = "TB" Then
      SubQOpenRecordset ResNF, _
        "SELECT * FROM NotasFiscais " & _
        "WHERE nofi_tx_controle = '" & XControle & "'  and empr_cd_empresa = " & PCodEmpresa & "", _
        "Estatico"
      ResNF.MoveFirst
      Do While Not ResNF.EOF
        'Atualiza a tabela de Duplicatas
        Conexao.Execute _
          "UPDATE Duplicatas SET " & _
          "  dupl_tx_docpagrec = '" & XFT_DOCPAGREC & "'," & _
          "  dupl_tx_status = 'R'," & _
          "  dupl_dt_pagrec = " & FunNuloData(DtpDtPagto.Value, NomeSgbd) & " " & _
          "WHERE dupl_nr_fatura = " & ResNF!nofi_nr_fatura
        ResNF.MoveNext
      Loop
      ResNF.Close
      'Atualiza a tabela de NotasFiscais
      Conexao.Execute _
        "UPDATE NotasFiscais SET " & _
        "  nofi_tx_status = '" & StatusDupl & "' " & _
        "WHERE nofi_tx_controle = '" & XControle & "'  and empr_cd_empresa = " & PCodEmpresa & ""
    End If
  End If
  
  '***************CONTABILIDADE*******************
'  SubContabilizarParcela Adodc1.Recordset.Fields("dupl_cd_Duplicata"), DtpDtPagto.Value, False, False
  '***********************************************
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
    
    If MsgBox("Deseja cancelar o registro?", vbInformation + vbYesNo, "CUIDADO") = vbYes Then
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
    
        'Muda o status da duplicata para cancelado
        'Alterado em 10/03/09 - Larissa
        Conexao.Execute "UPDATE Duplicatas SET dupl_tx_Status = 'X' WHERE dupl_cd_duplicata = " & Adodc1.Recordset.Fields("dupl_cd_duplicata")
        
        'Verifica se é a última duplicata da nota fiscal. Se for mudar o status da nota para realizado
        'verifica se todas as duplicatas são canceladas, se sim, muda o status da nota para cancelado
        XLT_SQL = "SELECT DISTINCT dupl_tx_status FROM Duplicatas WHERE dupl_nr_fatura = " & Adodc1.Recordset.Fields("dupl_nr_fatura")
        SubQOpenRecordset XLO_CHECK, XLT_SQL, Estatico
        
        If Not XLO_CHECK.EOF Then
            
            If XLO_CHECK.RecordCount = 1 And XLO_CHECK!dupl_tx_status = "X" Then
                
                Conexao.Execute "UPDATE NotasFiscais SET nofi_tx_Status = 'X' WHERE nofi_nr_fatura = " & Adodc1.Recordset.Fields("dupl_nr_fatura")
            
            Else
                
                XLO_CHECK.Filter = "dupl_tx_status='C'"
                
                If XLO_CHECK.EOF Then
                        
                        Conexao.Execute "UPDATE NotasFiscais SET nofi_tx_Status = 'R' WHERE nofi_nr_fatura = " & Adodc1.Recordset.Fields("dupl_nr_fatura")
                
                End If
                
            End If
        
        End If
                    
        XLO_CHECK.Close
        
      
         If XGT_TIPOEXPORTACAO = "PC" Then 'Cadastro do cancelamento da provisão. Apenas para Base Provisão - Caixa
                    'Cadastrar na contabilidade o inverso da provisão.
                    'Abrir a tabela de lançamentos
                    XLT_SQL = "SELECT lancamentos.*, notasfiscais.nofi_nr_documento FROM lancamentos LEFT JOIN notasfiscais ON lancamentos.lanc_cd_origemfin = notasfiscais.nofi_cd_notafiscal " & _
                                "WHERE (lanc_tx_Origem = 'FN' or lanc_tx_Origem = 'T') AND nofi_nr_fatura = " & Adodc1.Recordset.Fields("dupl_nr_fatura") & ""
                    SubQOpenRecordset XLO_CHECK, XLT_SQL, Estatico
                   
                    If Not (XLO_CHECK.EOF) Then '22/10/2009 Patrícia. Registra na contabilidade se for Base - Provisão e Caixa. Se for Base Caixa, o cancelamento não é realiado.
                            'Verifica se é duplicata a pagar ou a receber para pegar o valor
                            XLT_HISTORICO = "CANCELAMENTO DA DUPLICATA: " & Adodc1.Recordset.Fields("dupl_nr_fatura") & "-" & Adodc1.Recordset.Fields("dupl_nr_duplicata") & " DOC: " & XLO_CHECK!nofi_nr_documento & " - " & Adodc1.Recordset.Fields("FOCL_TX_FANTASIA")
            
                            While Not XLO_CHECK.EOF
                                XLI_CODLANC = funDefinirCodigoLanc(Format(Me.dtpCancelamento.Value, "dd/MM/yyyy"), XLI_LOTE)
                                 
                                XLT_SQL = "INSERT INTO Lancamentos(lanc_dt_data,lanc_nr_Lote,lanc_cd_Lanc,empr_cd_empresa," & _
                                                    "plco_cd_Debito,cere_cd_PcrDebitar,plco_cd_Credito,cere_cd_PcrCreditar,lanc_vl_valor," & _
                                                    "lanc_tx_Historico,lanc_tx_DescCustos,iden_cd_Projeto,lanc_cd_OrigemFin," & _
                                                    "lanc_tx_Origem,lanc_cd_estoque,lanc_cd_Patrimonio)" & _
                                            " VALUES (" & FunNuloData(Format(Me.dtpCancelamento.Value, "dd/MM/yyyy"), NomeSgbd) & ", " & XLI_LOTE & ", " & XLI_CODLANC & ", " & PCodEmpresa & "," & _
                                                    "" & FunNuloBanco(XLO_CHECK!plco_cd_credito) & "," & FunNuloBanco(XLO_CHECK!cere_cd_pcrCreditar) & "," & FunNuloBanco(XLO_CHECK!plco_cd_Debito) & "," & FunNuloBanco(XLO_CHECK!cere_cd_pcrDebitar) & "," & FunTrataFloat(Me.Adodc1.Recordset("valor")) & "," & _
                                                    "'" & XLT_HISTORICO & "','" & XLT_HISTORICO & "', " & FunNuloBanco(XLO_CHECK!iden_cd_Projeto) & ", " & Adodc1.Recordset.Fields("dupl_cd_duplicata") & "," & _
                                                    "'FD', " & FunNuloBanco(XLO_CHECK!lanc_cd_estoque) & ", " & FunNuloBanco(XLO_CHECK!lanc_cd_Patrimonio) & ")"
                                Conexao.Execute XLT_SQL
                                 
                                XLO_CHECK.MoveNext
                            Wend
                        
                            XLO_CHECK.Close
                            Set XLO_CHECK = Nothing
                    Else
                            MsgBox "Não existe registro de Provisão na contabilidade para o documento selecionado. O cancelamento será feito apenas no financeiro", vbInformation
                    End If
         End If
        
        Conexao.CommitTrans
        
        
        '****************REGISTRA LOG DA OPERAÇÃO***********************
          ReDim XGM_MATRIZLOG(3, 2) As Variant
          XGM_MATRIZLOG(0, 0) = "Parcela" 'Campo de identificação da duplicata.
          XGM_MATRIZLOG(1, 0) = "Nº Doc" 'Número de documento.
          XGM_MATRIZLOG(2, 0) = "Valor" 'Campo fatura da duplicata.
          XGM_MATRIZLOG(0, 1) = Adodc1.Recordset.Fields("Parcela") 'Pega a identificação da duplicata.
          XGM_MATRIZLOG(1, 1) = Adodc1.Recordset.Fields("Numdocumento") 'Pega número de documento da nota fiscal que a parcela pertence.
          XGM_MATRIZLOG(2, 1) = Adodc1.Recordset.Fields("Valor") 'Pega o valor da duplicata.
                    
          Call subRegistraLog("TelaDuplicatas", "1", funCriaDescricaoLog(XGM_MATRIZLOG, CANCELAR, DUPLICATA))
          
        ' *************************************************************
        
        MsgBox ("Cancelamento realizado c/ sucesso !")
        
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    
    End If
    Me.panCancelar.Visible = False
    Exit Sub
    
TrataErro:
    funTrataErros (ComMensagem)
    Conexao.RollbackTrans
    Exit Sub
   
    
End Sub

Private Sub cmdConfirmarCartaTransferencia_Click()
    Dim XLT_CIDADE As String
    Dim XLT_HISTORICO As String
    Dim XLT_EXTENSO As String
    Dim XLT_TEMDESC As String
    Dim XLD_VALOR As Double
    Dim XLF_VLBRUTO As Double
    Dim XLF_VALORGLOSA As Double
    Dim XLO_SUBRELATORIO As ADODB.Recordset
    Dim XLO_SUBRELATORIO2 As ADODB.Recordset
    Dim XLO_SUBRELATORIO3 As ADODB.Recordset
    Dim XLO_RSEMPRESA     As ADODB.Recordset
    Dim XLO_NOTA As ADODB.Recordset
    Dim XLO_RSCCDESTINO     As ADODB.Recordset '16/10/08-Patrícia
    Dim XLO_DUPLICATA As Long '16/10/08-Patrícia
    
    
    XFT_TEXTO = TxtTextoRecibo.Text
    If Modelo = 0 Then '30/09/08 - Patrícia
        subImprimeRecibo (Trim(txtNumeroRecibo.Text))
        txtNumeroRecibo.Text = ""
        txtEmissor.Text = ""
        Exit Sub
    End If
    'Verifica se foi informado o número do recibo
    If Trim$(txtNumeroRecibo.Text) = "" Then
        MsgBox "É preciso informar o número da carta!", vbInformation, "Campo obrigatório"
        txtNumeroRecibo.SetFocus
        Exit Sub
    End If
    
    If optModelo1.Value = True Then '02/10/08 - Patrícia
       Modelo = 1
    ElseIf optModelo2.Value = True Then
       Modelo = 2
    Else
       Modelo = 3
    End If
    
    TelaDuplicatas.MousePointer = vbHourglass
    
    SubQOpenRecordset XLO_NOTA, "SELECT nofi_tx_ReferenciaTransferencia, focl_cd_Forncli FROM NotasFiscais WHERE nofi_nr_Fatura = " & Adodc1.Recordset.Fields("dupl_nr_Fatura"), Estatico
    
    'Pega o histórico da transferência
    XLT_HISTORICO = FunNulo(XLO_NOTA!nofi_tx_ReferenciaTransferencia)
    
    'Se a duplicata não tiver fornecedor é porque é uma transferência bancária.
    If Not (IsNull(XLO_NOTA!focl_cd_forncli)) Then
        Set XLO_SUBRELATORIO3 = ConexaoRelatorio.Execute("SELECT * FROM FornClientes WHERE focl_cd_Forncli = " & XLO_NOTA!focl_cd_forncli & "")
    Else
        Set XLO_SUBRELATORIO3 = ConexaoRelatorio.Execute("SELECT empr_tx_razaoSocial FROM Empresas WHERE empr_cd_empresa = " + CStr(PCodEmpresa))
    End If
    
    XLO_NOTA.Close
    
    XLT_CIDADE = "Salvador"
    
    
    XLD_VALOR = CStr(TDBGrid1.Columns("A Receber"))
    XLT_EXTENSO = funExtenso(TDBGrid1.Columns("A Receber"))
    
    FrmVisRelatoriosEXT.Caption = "Carta de Transferência"
 
    'Atualiza o número de controle da duplicata
    Conexao.Execute "UPDATE duplicatas SET dupl_tx_controle = '" & txtNumeroRecibo.Text & "' WHERE dupl_cd_duplicata = " & Adodc1.Recordset.Fields("dupl_cd_duplicata")
        
    'Pega a duplicata do tipo credito que foi gerada na transferência -16/10/08 -Patrícia
    SubQOpenRecordset XLO_RSCCDESTINO, "SELECT dupl_cd_duplicata FROM ConsFINTransfBancaria WHERE nofi_tx_controle = " & Adodc1.Recordset.Fields("nofi_tx_controle") & " AND nofi_tx_TipoMov = 'S'", Estatico
    XLO_DUPLICATA = XLO_RSCCDESTINO!dupl_cd_Duplicata
    XLO_RSCCDESTINO.Close
    
    Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " + CStr(PCodEmpresa))
    
    Select Case Modelo '30/09/08 - Patrícia
    Case Is = 1
        Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDuplicatas WHERE dupl_cd_duplicata = " & XLO_DUPLICATA & "")
        Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\CartaTransferenciaBancariaModelo1.rpt")
    Case Is = 2
        Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDuplicatas WHERE dupl_cd_duplicata = " & Adodc1.Recordset.Fields("dupl_cd_duplicata") & "")
        Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\CartaTransferenciaBancariaModelo2.rpt")
    Case Is = 3
            Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDuplicatas WHERE dupl_cd_duplicata = " & Adodc1.Recordset.Fields("dupl_cd_duplicata") & "")
            Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\CartaTransferenciaBancariaModelo3.rpt")
    End Select
       
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO

    'Ativa os Relatórios
    With XGR_RELATORIO.FormulaFields
        .GetItemByName("Recibo").Text = Chr(34) + txtNumeroRecibo + Chr(34)
        .GetItemByName("Empresa").Text = Chr(34) + CStr(PCodEmpresa) + Chr(34)
        .GetItemByName("NomeEmpresa").Text = Chr(34) + CStr(PEmpresa) + Chr(34)
        .GetItemByName("Valor").Text = Chr(34) + CStr(Format(XLD_VALOR, "standard")) + Chr(34)
        .GetItemByName("ValorExtenso").Text = Chr(34) + XLT_EXTENSO + Chr(34)
        .GetItemByName("Cidade").Text = Chr(34) + XLT_CIDADE + Chr(34)
        .GetItemByName("Historico").Text = Chr(34) & FunTiraAspas(XLT_HISTORICO) & Chr(34) '29/09/08 - Patrícia
        .GetItemByName("DataEmissaoCarta").Text = Chr(34) & Day(DtpDataCartaTransferencia.Value) & " de " & MonthName(Month(DtpDataCartaTransferencia.Value), False) & " de " & Year(DtpDataCartaTransferencia.Value) & Chr(34)
    End With
    
    
    Select Case Modelo ' Select inserido em 02/10/2008
    Case Is = 1
        With XGR_RELATORIO.FormulaFields
            .GetItemByName("Emissor").Text = Chr(34) + Trim(txtEmissor) + Chr(34)
            .GetItemByName("HistoricoCarta").Text = Chr(34) + Trim(XLT_HISTORICO) + Chr(34)
            .GetItemByName("CnpjEmpresa").Text = Chr(34) & "CNPJ: " & Trim(XGO_RSCABECALHO!empr_tx_CGC) & Chr(34)
        End With
        
        If (Adodc1.Recordset.Fields("dupl_tx_debcred")) <> "C" Then ' 15/10/08 - Verifica se a duplicata é do tipo debito. Se for, procura a duplicata do tipo credito.
                'Procura a dulpicata do tipo crédito
                SubQOpenRecordset XLO_RSCCDESTINO, "SELECT dupl_cd_duplicata FROM ConsFINTransfBancaria WHERE nofi_tx_controle = " & Adodc1.Recordset.Fields("nofi_tx_controle") & " AND nofi_tx_TipoMov = 'E'", Estatico
                XLO_DUPLICATA = XLO_RSCCDESTINO!dupl_cd_Duplicata
                XLO_RSCCDESTINO.Close
                
                'Pega os dados da conta na duplicata do tipo crédito
                SubQOpenRecordset XLO_RSCCDESTINO, "SELECT * FROM ConsFINDuplicatas WHERE dupl_cd_duplicata = " & XLO_DUPLICATA & " AND nofi_tx_TipoMov = 'E'", Estatico
                        With XGR_RELATORIO.FormulaFields
                        .GetItemByName("Coco_cd_agencia").Text = Chr(34) + XLO_RSCCDESTINO!coco_cd_agencia + Chr(34)
                        .GetItemByName("Coco_nr_dagtransf").Text = Chr(34) + XLO_RSCCDESTINO!coco_nr_dag + Chr(34)
                        .GetItemByName("Coco_nm_agencia").Text = Chr(34) + XLO_RSCCDESTINO!coco_nm_agencia + Chr(34)
                        .GetItemByName("Coco_tx_conta").Text = Chr(34) + XLO_RSCCDESTINO!coco_tx_conta + Chr(34)
                        .GetItemByName("Coco_tx_descricao").Text = Chr(34) + XLO_RSCCDESTINO!coco_tx_descricao + Chr(34)
                        .GetItemByName("Coco_nr_dactransf").Text = Chr(34) + XLO_RSCCDESTINO!coco_nr_dac + Chr(34)
                        End With
                        XLO_RSCCDESTINO.Close
        Else
                        With XGR_RELATORIO.FormulaFields
                        .GetItemByName("Coco_cd_agencia").Text = Chr(34) + Adodc1.Recordset.Fields("coco_cd_agencia") + Chr(34)
                        .GetItemByName("Coco_nr_dag_transf").Text = Chr(34) + Adodc1.Recordset.Fields("coco_nr_dag") + Chr(34)
                        .GetItemByName("Coco_nm_agencia").Text = Chr(34) + Adodc1.Recordset.Fields("coco_nm_agencia") + Chr(34)
                        .GetItemByName("Coco_tx_conta").Text = Chr(34) + Adodc1.Recordset.Fields("coco_tx_conta") + Chr(34)
                        .GetItemByName("Coco_tx_descricao").Text = Chr(34) + Adodc1.Recordset.Fields("Coco_tx_descricao") + Chr(34)
                        .GetItemByName("Coco_nr_dac_transf").Text = Chr(34) + Adodc1.Recordset.Fields("coco_nr_dac") + Chr(34)
                        End With
        End If

    Case Is = 3
            With XGR_RELATORIO.FormulaFields
            .GetItemByName("CnpjEmpresa").Text = Chr(34) & "CNPJ: " & Trim(XGO_RSCABECALHO!empr_tx_CGC) & Chr(34)
            End With

    Case Is = 2
            With XGR_RELATORIO.FormulaFields
            .GetItemByName("CnpjEmpresa").Text = Chr(34) & "CNPJ: " & Trim(XGO_RSCABECALHO!empr_tx_CGC) & Chr(34)
            End With
    End Select
    
    
    FrmVisRelatoriosEXT.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
    FrmVisRelatoriosEXT.SubVerRelatorio XGR_RELATORIO, ""
    
    TelaDuplicatas.MousePointer = vbDefault
            
    txtNumeroRecibo.Text = ""
    txtEmissor.Text = ""

End Sub

Private Sub CmdConfirmaRecibo_Click()
    Dim ResTipoDoc As ADODB.Recordset
    Dim XLT_SQL As String
    Dim XLO_RSEMPRESA As ADODB.Recordset
    
    XFT_TEXTO = TxtTextoRecibo.Text
    
    SubQOpenRecordset XLO_RSEMPRESA, "SELECT empr_tx_cliente FROM Empresas WHERE empr_cd_empresa = " + CStr(PCodEmpresa), "Estatico"
    If Not (XLO_RSEMPRESA.EOF) Then ' If inserido em 27/10/08 - Patrícia. Inclusão de relatórios solicitados pelo cliente Instituto Aliança
        If Trim(XLO_RSEMPRESA!empr_tx_cliente) = "ALIANCA" Then
                PanDetalhesCartaTransferencia.Visible = True
        Else
                XLT_SQL = "SELECT * FROM Duplicatas WHERE dupl_cd_duplicata = " & Adodc1.Recordset.Fields("dupl_cd_duplicata")
                SubQOpenRecordset ResTipoDoc, XLT_SQL, Dinamico
                If IsNull(ResTipoDoc!dupl_tx_controle) Then
                    'Emitir o recibo e gera o nº de controle
                    If MsgBox("Confirma emissão do Recibo?", vbQuestion + vbYesNo, "ATENÇÃO") = vbYes Then
                        If Adodc1.Recordset.Fields("dupl_tx_DebCred") = "D" Then  'D ou C
                            ResTipoDoc!dupl_tx_controle = CStr(FunPegaGuardaUltimo("C"))
                        Else
                            ResTipoDoc!dupl_tx_controle = CStr(FunPegaGuardaUltimo("R"))
                        End If
                        subImprimeRecibo (ResTipoDoc!dupl_tx_controle)
                    Else
                        Exit Sub
                    End If
                Else
                    If MsgBox("O recibo já foi emitido. Deseja reemitir ?", vbYesNo + vbCritical, "ATENÇÃO") = vbYes Then
                        'Reemitir sem gerar um novo número de controle
                        subImprimeRecibo (ResTipoDoc!dupl_tx_controle)
                    Else
                        Exit Sub
                    End If
                End If
                ResTipoDoc.Update
                ResTipoDoc.Close

        End If
    End If
    XLO_RSEMPRESA.Close
    
    PanTextoRecibo.Visible = False
End Sub

Private Sub CmdConfirmaReciboAlianca_Click()
    Dim ResTipoDoc As ADODB.Recordset
    Dim XLT_SQL As String
    Dim XLO_NUMREC As ADODB.Recordset
    
    XFT_TEXTO = TxtTextoRecibo.Text
    
    XLT_SQL = "SELECT * FROM Duplicatas WHERE dupl_cd_duplicata = " & Adodc1.Recordset.Fields("dupl_cd_duplicata")
    SubQOpenRecordset ResTipoDoc, XLT_SQL, Dinamico
    If IsNull(ResTipoDoc!dupl_tx_controle) Then
        XLT_SQL = "SELECT * FROM Duplicatas WHERE UPPER(dupl_tx_controle) = '" & Trim(UCase(txtNumeroRecibo.Text)) & "'"
        SubQOpenRecordset XLO_NUMREC, XLT_SQL, Dinamico
        If Not (XLO_NUMREC.EOF) Then
           MsgBox "O número de recibo informado já está em uso.", vbCritical
           txtNumeroRecibo.SetFocus
           Exit Sub
        End If
        'Emitir o recibo com o número de ocntrole informado pelo usuário
        If MsgBox("Confirma emissão do Recibo?", vbQuestion + vbOKCancel, "ATENÇÃO") = vbOK Then
           ResTipoDoc!dupl_tx_controle = Trim(txtNumeroRecibo.Text)
           subImprimeRecibo (ResTipoDoc!dupl_tx_controle)
        Else
          PanTextoRecibo.Visible = False
          txtNumeroRecibo.Text = ""
          Exit Sub
        End If

    Else
        If MsgBox("O recibo já foi emitido. Deseja reemitir ?", vbYesNo + vbCritical, "ATENÇÃO") = vbYes Then
            'Reemitir sem gerar um novo número de controle
            subImprimeRecibo (ResTipoDoc!dupl_tx_controle)
        Else
            PanTextoRecibo.Visible = False
            txtNumeroRecibo.Text = ""
            Exit Sub
        End If
    End If
    ResTipoDoc.Update
    ResTipoDoc.Close
    
    PanDetalhesCartaTransferencia.Visible = False
    txtNumeroRecibo.Text = ""

End Sub

Private Sub CmdCopia_Click()
    If FunTabelaVazia(Adodc1) Then
        MsgBox "Não pode ser emitida nenhuma cópia, pois a tabela está vazia", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    
    'Não permitir efetivação para registros de Estorno Financeiro
    If Left(Adodc1.Recordset.Fields("Numdocumento"), 1) = "E" Then
        MsgBox "Esta opção não é permitida para registro de Estorno Financeiro!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    
    'Não permitir efetivação para transferências de custo:
    If TDBGrid1.Columns("Tipo Doc.") = "TRANSFERÊNCIA DE CUSTO" Then
        MsgBox "Esta opção não é permitida para um registro de transferência de custo!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    
    TDBGrid1.Enabled = False
    subDesabilitaBotoes
    CmdSair.Enabled = False
        
    PanCopia.Left = (TDBGrid1.Width - PanCopia.Width) / 2
    PanCopia.Top = (TDBGrid1.Height - PanCopia.Height) / 2
    PanCopia.Visible = True
    PanCopia.Enabled = True
    
    TxtNroCheque2.Text = TDBGrid1.Columns("Nº Doc. Pagto./Rec.").Value
    CboCCorrente2.BoundText = Adodc1.Recordset.Fields("coco_cd_codigo")
    
    TxtNroCheque2.SetFocus
End Sub

Private Sub CmdDescAcresc_Click()
    Dim ResNF As ADODB.Recordset
    
    'Os registros frutos de um estorno não devem atender a nenhuma solicitação
    If Adodc1.Recordset.Fields("dupl_tx_Estorno") = "S" Then
        MsgBox "Não é possível cadatrar um Desconto/Acréscimo para um estorno!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    
    'Não permitir efetivação para registros de Estorno Financeiro
    If Left(Adodc1.Recordset.Fields("Numdocumento"), 1) = "E" Then
        MsgBox "Esta opção não é permitida para registro de Estorno Financeiro!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    
    'Não permitir efetivação para transferências de custo:
    If TDBGrid1.Columns("Tipo Doc.") = "TRANSFERÊNCIA DE CUSTO" Then
        MsgBox "Esta opção não é permitida para um registro de transferência de custo!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    
    Chave = Adodc1.Recordset.Fields("dupl_cd_duplicata")
    If Not IsNull(Adodc1.Recordset.Fields("Valor")) Then
        XVlDupl = Adodc1.Recordset.Fields("Valor")
    Else
        XVlDupl = 0
    End If
    xChave = Adodc1.Recordset.Bookmark
'    TelaDescAcresc.Show 1
    
    Set Formulario = TelaDuplicatas
    TDBGrid1.Refresh

    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    Adodc1.Recordset.Bookmark = xChave
    
End Sub

Private Sub CmdEstorno_Click()
    Dim Sql As String
    Dim codbancoforn As Integer
    Dim XLO_LANC As ADODB.Recordset
    Dim XLL_CODESTORNO As Long
    Dim XLB_CONTABILIZOU As Boolean
    Dim XLT_TIPO As String
    Dim XLO_DESCACRES As ADODB.Recordset
    
    
    If Not FunVerifica_Permissao("TelaDuplicatas", "1") Then
        Exit Sub
    End If
    
    'Não permitir efetivação para registros de Estorno Financeiro
    If Left(Adodc1.Recordset.Fields("Numdocumento"), 1) = "E" Then
        MsgBox "Esta opção não é permitida para registro de Estorno Financeiro!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    
    'Não permitir efetivação para transferências de custo:
    If TDBGrid1.Columns("Tipo Doc.") = "TRANSFERÊNCIA DE CUSTO" Then
        MsgBox "Esta opção não é permitida para um registro de transferência de custo!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    
    XLB_CONTABILIZOU = False
       
    Posicao = Adodc1.Recordset.Bookmark
    If TDBGrid1.Columns("status").Text = "Realizado" And TDBGrid1.Columns("Tipo Pagto./Rec.").Text <> "Dinheiro" Then
        '***************CONTABILIDADE*******************
        'possui o módulo de contabilidade
        If XGB_PQUALICONT Then
            If TDBGrid1.Columns("Tipo Pagto./Rec.").Text <> "Crédito em Conta" Then
                If funMesFechado(Adodc1.Recordset.Fields("dupl_dt_PagRec")) <> "A" Then
                    If Adodc1.Recordset.Fields("dupl_tx_debcred") = "D" Then
                        MsgBox "Não é possível cancelar o Pagamento, pois o mês contábil já foi fechado. " _
                            + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
                    Else
                        MsgBox "Não é possível cancelar o Recebimento nessa data, pois o mês contábil já foi fechado. " _
                            + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
                    End If
                    Exit Sub
                End If
            End If
        End If
        '***********************************************
        
        If Adodc1.Recordset.Fields("dupl_tx_Estorno") = "S" Then
            If MsgBox("Parcela já estornada! Deseja cancelar o estorno?", vbYesNo + vbCritical, "ATENÇÃO") = vbYes Then
               'Deletar o Estorno da tabela Lançamentos
               '***************CONTABILIDADE*******************
               If (Adodc1.Recordset.Fields("dupl_tx_debcred") = "D" And Adodc1.Recordset.Fields("Saida") < 0) Or (Adodc1.Recordset.Fields("dupl_tx_debcred") = "E" And Adodc1.Recordset.Fields("Entrada") < 0) Then
                   'O código gravado na contabilidade é sempre o positivo.Apagar o código+1
                   XLL_CODESTORNO = Adodc1.Recordset.Fields("dupl_cd_duplicata") + 1
               Else
                   XLL_CODESTORNO = Adodc1.Recordset.Fields("dupl_cd_duplicata")
               End If
              
                Conexao.BeginTrans
               
               '/*/*/*/**VERIFICAR SE EXISTEM LANÇAMENTOS CONTÁBEIS DO PAGTO/RECEB**/*/*/*/
               'só para registros <> de dinheiro e cheque
               XLT_TIPO = Adodc1.Recordset.Fields("dupl_tx_tipopagrec")
               If XLT_TIPO <> "6" And XLT_TIPO <> "1" Then
                     '(13/11/06) - Inclusão de dupl_tx_estorno <> 'S' para ser revisado
                     Sql = "SELECT * FROM Duplicatas " & _
                              "WHERE dupl_tx_DocPagRec = '" & Adodc1.Recordset.Fields("dupl_tx_DocPagRec") & "' " & _
                                     " AND coco_cd_codigo = " & Adodc1.Recordset.Fields("coco_cd_codigo") & " " & _
                                     " AND dupl_dt_PagRec = " & FunNuloData(Adodc1.Recordset.Fields("dupl_dt_PagRec"), NomeSgbd) & " " & _
                                     " AND dupl_nr_Fatura <> " & Adodc1.Recordset.Fields("dupl_nr_Fatura") & "" & _
                                     " AND dupl_tx_estorno <> 'S' " & _
                                     " AND dupl_tx_TipoPagRec = '" & Adodc1.Recordset.Fields("dupl_tx_TipoPagRec") & "' "
                     SubQOpenRecordset XLO_LANC, Sql, Estatico
                       
                    
                    If Not XLO_LANC.EOF Then
                        'qdo existir mais de um no pagto consolidado
                        MsgBox "Cancele o pag/rec original antes de cancelar o estorno.", vbCritical, "ATENÇÃO"
                        Conexao.CommitTrans
                        Exit Sub
                     Else
                        'Qdo só existir um titulo e este for D/C, C/C ou borderô
                        Sql = "DELETE FROM Lancamentos " & _
                                "WHERE lanc_cd_OrigemFin = " & XLL_CODESTORNO & " " & _
                                    " AND lanc_tx_Origem = 'FD' " & _
                                    " AND empr_cd_empresa = " & PCodEmpresa
                        Conexao.Execute Sql
                    End If
               End If
               '/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/
               
                'Excluindo Desc/Acresc das duplicatas estornadas
               If TDBGrid1.Columns("Desconto/Acréscimo").Text <> "" Then
                    XGT_SQL = ""
                    XGT_SQL = "SELECT dupl_cd_duplicata FROM Duplicatas " & _
                            " WHERE dupl_nr_fatura = " & CStr(TDBGrid1.Columns("Fatura")) & " And dupl_tx_Estorno = 'S'"
                    SubQOpenRecordset XLO_DESCACRES, XGT_SQL, Estatico
                                                                
                    While Not XLO_DESCACRES.EOF
                        Sql = "DELETE FROM DuplDescAcresc WHERE dupl_cd_duplicata = " & "'" & XLO_DESCACRES!dupl_cd_Duplicata & "'"
                        Conexao.Execute Sql
                        XLO_DESCACRES.MoveNext
                    Wend
                          XLO_DESCACRES.Close
                End If
                    
               Conexao.Execute "DELETE FROM Duplicatas WHERE dupl_nr_fatura = " & CStr(TDBGrid1.Columns("Fatura")) & " And dupl_tx_Estorno = 'S'"
       
               Sql = "DELETE FROM Lancamentos " & _
                         "WHERE lanc_cd_OrigemFin = " & XLL_CODESTORNO & " " & _
                                " AND lanc_tx_Origem = 'FE' " & _
                                " AND empr_cd_empresa = " & PCodEmpresa
               Conexao.Execute Sql
               
               'Excluir os lançamentos contábeis quando o estorno for um cheque
                
                XGT_SQL = "SELECT dupl_cd_duplicata, dupl_tx_TipoPagRec FROM Duplicatas " & _
                       " WHERE dupl_nr_fatura = " & CStr(TDBGrid1.Columns("Fatura")) & " And dupl_tx_Estorno is null"
                SubQOpenRecordset XLO_DESCACRES, XGT_SQL, Estatico
                
                If XLO_DESCACRES!dupl_tx_tipopagrec = 1 Then
                     Sql = "DELETE FROM Lancamentos " & _
                              "WHERE lanc_cd_OrigemFin = " & XLO_DESCACRES!dupl_cd_Duplicata & " " & _
                                     " AND lanc_tx_Origem = 'FD' " & _
                                     " AND empr_cd_empresa = " & PCodEmpresa
                    Conexao.Execute Sql
                End If
                
                XLO_DESCACRES.Close
               
               
               '***********************************************
               Conexao.CommitTrans
               
               subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
            Else
               Exit Sub
            End If
        Else
            If MsgBox("Confirma o estorno da parcela?", vbYesNo + vbCritical, "ATENÇÃO") = vbYes Then
                PanEstorno.Left = (TDBGrid1.Width - PanEstorno.Width) / 2
                PanEstorno.Top = (TDBGrid1.Height - PanEstorno.Height) / 2
                DtpDtEstorno.Value = TDBGrid1.Columns("Pagto./Rec.")
                PanEstorno.Enabled = True
                PanEstorno.Visible = True
                
            Else
                Exit Sub
            End If
        End If
    ElseIf TDBGrid1.Columns("status").Text <> "Realizado" Then
        MsgBox "Não é possível estornar uma parcela cujo satus não seja 'Realizado'!", vbCritical, "ATENÇÃO!"
    ElseIf TDBGrid1.Columns("Tipo Pagto./Rec.").Text = "Dinheiro" Then
        MsgBox "Não é possível estornar uma parcela cujo Tipo de Pagto./Rec. seja 'Dinheiro'!", vbCritical, "ATENÇÃO!"
    End If

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
    
    Screen.MousePointer = VBRUN.MousePointerConstants.vbHourglass
        
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

    DoEvents
   
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

Private Sub CmdImpCheque_Click()
    On Error GoTo TrataErro:
    Dim GuardY, GuardX As Single
    Dim TamanhoPapel As Single
    Dim Campo As String
    Dim LargPapel As Single, AltPapel As Single
    Dim ResCheque As ADODB.Recordset
    Dim Sql As String
    Dim XLO_IMPRESSORA As Printer
    Dim XLT_IMPRESSORA As String
        
    Dim VlTotal As Double
    Dim XCodBanco As Integer 'Armazena o código do banco

    XCodBanco = DatContaCorrente.Recordset.Fields("banc_cd_codigo")
    
    If Adodc1.Recordset.Fields("da") <> "" Then
        VlTotal = CDbl(Adodc1.Recordset.Fields("Valor")) + CDbl(Adodc1.Recordset.Fields("da"))
    Else
        VlTotal = CDbl(Adodc1.Recordset.Fields("Valor"))
    End If
   
       
    'Nome da impressora padrão
    XLT_IMPRESSORA = Printer.DeviceName
    FrmImpressora.Show vbModal
    
   
    If MsgBox("Prepare o formulário e confirme a impressão.", vbQuestion + vbYesNo, "ATENÇÃO") = vbNo Then Exit Sub
    
    Printer.ScaleMode = 7  'vbCentimeters
    
    Printer.Orientation = 1  'Portrait
    
    TamanhoPapel = Printer.PaperSize
    LargPapel = 17.5 / 567 'Convertendo Twips para Cms
    AltPapel = 7.6 / 567 'Convertendo Twips para Cms
    
    'ImprimeCheque "Cheque", Formulario
'    funImprimeCheque Formulario, VlTotal, XCodBanco, FunNulo(TxtCidade.Text), FunNulo(CboNominal.Text), Format(DtpDtPagto.Value, "dd/mm/yyyy")
    
    'Insere a descricao nominal e a descricao verso na tabela VersoCheques
    If TxtVerso.Text <> "" Then
        If CboCCorrente.Text <> "" Then
            Conexao.Execute "INSERT INTO VersoCheques(coco_cd_codigo,vech_tx_numcheque,vech_tx_descricaonominal,vech_tx_descricaoverso) " & _
            " VALUES (" & CboCCorrente.BoundText & ",'" & TxtNroCheque.Text & "','" & FunTrataString(FunNulo(CboNominal.Text)) & "','" & FunTrataString(TxtVerso.Text) & "')"
        Else ' BAIXA PARCIAL
            Conexao.Execute "INSERT INTO VersoCheques(coco_cd_codigo,vech_tx_numcheque,vech_tx_descricaonominal,vech_tx_descricaoverso) " & _
            " VALUES (" & CboCCorrenteParcial.BoundText & ",'" & TxtNroCheque.Text & "','" & FunTrataString(FunNulo(CboNominal.Text)) & "','" & FunTrataString(TxtVerso.Text) & "')"
        End If
    End If
    
    Printer.PaperSize = TamanhoPapel
    
    If MsgBox("Confirma emissão da cópia do cheque?", vbQuestion + vbYesNo, "ATENÇÃO") = vbNo Then
        Exit Sub
    Else
        subImprimeCopia
    End If

       
    'Retorna para a impressora padrão
    For Each XLO_IMPRESSORA In Printers
      If XLO_IMPRESSORA.DeviceName = XLT_IMPRESSORA Then
        Set Printer = XLO_IMPRESSORA
        Exit For
      End If
    Next
    
    CmdRetorna_Click
    
Exit Sub

TrataErro:
    funTrataErros (ComMensagem)
End Sub

Private Sub CmdImpCopia_Click()
  Dim VlTotal As Double, vlpercent As Double
  Dim xnota As String, Sql As String
  Dim total As String, Banco As String, nroChCons As String
  Dim XNominal As String, Data As String, ProxCheque As String
  Dim NroCheque As String
  Dim xChave As Variant
  Dim Posicao As Integer, Cont As Integer
  Dim TemCondicao As Boolean, TemCondicaoCons As Boolean, ChSimples As Boolean, ChCons As Boolean
  Dim ResDistrib As ADODB.Recordset, ResCheques As ADODB.Recordset
  Dim XCidade As String
  Dim VlDistrib As String, percent As String
  Dim SelecaoCons, ConjuncaoCons As String
  Dim XLA_SUBRELATORIO As ADODB.Recordset
  Dim XLO_SUBRELATORIO2 As ADODB.Recordset
  Dim XLO_EMPRESA As ADODB.Recordset 'Utilizado para verificar o cliente que está executando - Larissa Moura 21/05/08

  
  XNominal = ""
  XCidade = "Salvador"
  
  'Verificações
  If TxtNroCheque2.Text = "" Then
    MsgBox "O nº do cheque é obrigatório!", vbCritical, "ATENÇÃO!"
    TxtNroCheque2.SetFocus
    Exit Sub
  End If
  If CboCCorrente2.Text = "" Then
    MsgBox "O campo Conta Corrente é obrigatório!", vbCritical, "ATENÇÃO!"
    CboCCorrente2.SetFocus
    Exit Sub
  End If
  If MsgBox("Confirma emissão da cópia do cheque?", vbQuestion + vbYesNo, "ATENÇÃO") = vbNo Then Exit Sub

'*** Pega o nome da empresa que está executando o sistema
'** Início - Larissa Moura 21/05/08
  SubQOpenRecordset XLO_EMPRESA, "SELECT TOP 1 empr_tx_cliente FROM empresas WHERE empr_tx_cliente IS NOT NULL", Estatico
'** Fim - Larissa Moura 21/05/08
  
  XGT_SELECAO = ""
  XGT_CONJUNCAO = ""
  SelecaoCons = ""
  ConjuncaoCons = ""
  XNominal = ""
  XCidade = "Salvador"
  xChave = TDBGrid1.Bookmark
  TelaDuplicatas.MousePointer = vbHourglass
  Cont = 0
  Sql = "select * from ConsFINCheques where (dupl_tx_debcred='D' and coco_cd_codigo = " & _
    CboCCorrente2.BoundText & " and dupl_tx_TipoPagRec = '1' and (dupl_tx_docpagrec >= '" & _
    TxtNroCheque2.Text & "' And dupl_tx_docpagrec <= '" & TxtNroCheque3.Text & _
    "') and (dupl_tx_Estorno<>'S' or (dupl_tx_estorno is null))) ORDER BY dupl_tx_docpagrec"
  
  SubQOpenRecordset ResCheques, Sql, "Estatico"
  While Not ResCheques.EOF
    NroCheque = ResCheques!dupl_tx_docpagrec
    
    If ResCheques.Bookmark = 1 Then
        XGT_SELECAO = "("
    End If
    
    If ResCheques.Bookmark = 1 Then
        SelecaoCons = "("
    End If
    
    ResCheques.MoveNext
    If Not ResCheques.EOF Then
      ProxCheque = ResCheques!dupl_tx_docpagrec
    Else
      ProxCheque = ""
    End If
    If Cont = 0 Then nroChCons = ""
    
    If NroCheque <> ProxCheque And NroCheque <> nroChCons Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "(ConsFINCheques.dupl_tx_docpagrec = '" & NroCheque & "')"
      XGT_CONJUNCAO = " OR "
      ChSimples = True
    Else
      If Cont = 0 Then
        nroChCons = NroCheque
        SelecaoCons = SelecaoCons & ConjuncaoCons & _
          "(ConsFINCopiaConsolidado.dupl_tx_docpagrec = '" & NroCheque & "')"
        
        ConjuncaoCons = " OR "
        'TemCondicaoCons = True
        ChCons = True
      End If
            
      If Not ResCheques.EOF Then ResCheques.MoveNext
      If Not ResCheques.EOF Then
        If NroCheque = ResCheques!dupl_tx_docpagrec Then
          Cont = Cont + 1
        Else
          Cont = 0
        End If
      End If
    End If
      
    If ResCheques.EOF Then
       XGT_SELECAO = XGT_SELECAO & ")"
    End If
    If ResCheques.EOF Then
       SelecaoCons = SelecaoCons & ")"
    End If
  Wend
  ResCheques.Close
  
  'Não existem cheques
  If Not ChSimples And Not ChCons Then
    MsgBox "Este nº de cheque não existe.", vbCritical, "ATENÇÃO!"
    TelaDuplicatas.MousePointer = vbDefault
    CmdRetornar2_Click
    Exit Sub
  End If
  
  'Existem cheques simples
  If ChSimples Then
    XGT_SELECAO = XGT_SELECAO + " and (ConsFINCheques.coco_cd_codigo = " & _
      CboCCorrente2.BoundText & ")"
    
    'Abre os recordsets e verifica se eles possuem dados
    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINCheques WHERE " & XGT_SELECAO)
    If XGO_RSRELATORIO.EOF Then
      MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
      Me.MousePointer = vbDefault
      Exit Sub
    End If
    Set XLA_SUBRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDistribCopiaCheque where (dupl_tx_debcred='D' and coco_cd_codigo = " & CboCCorrente2.BoundText & " and dupl_tx_TipoPagRec = '1' and (dupl_tx_docpagrec >= '" & TxtNroCheque2.Text & "' And dupl_tx_docpagrec <= '" & TxtNroCheque3.Text & "') and (dupl_tx_Estorno<>'S' or (dupl_tx_estorno is null))) ORDER BY dupl_tx_docpagrec")
    Set XLO_SUBRELATORIO2 = ConexaoRelatorio.Execute("SELECT * FROM ConsFINCChequeDescAcresc where (dupl_tx_debcred='D' and coco_cd_codigo = " & CboCCorrente2.BoundText & " and dupl_tx_TipoPagRec = '1' and (dupl_tx_docpagrec >= '" & TxtNroCheque2.Text & "' And dupl_tx_docpagrec <= '" & TxtNroCheque3.Text & "') and (dupl_tx_Estorno<>'S' or (dupl_tx_estorno is null))) ORDER BY dupl_tx_docpagrec")
    
    'Abre o relatório
    FrmVisRelatoriosEXT.Caption = "Cópia de Cheque Simples"
    
'*** Verifica qual empresa está executando o sistema para selecionar o relatório que ela utiliza
'** Início - Larissa Moura 21/05/08
    Select Case UCase(XLO_EMPRESA!empr_tx_cliente)
        Case "COSTA ANDRADE"
                Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\CopiaChequeCA.rpt")
        Case Else
                Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\CopiaCheque.rpt")
    End Select
'** Fim - Larissa Moura 21/05/08
    
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO

    'Imprime o relatório
    XGR_RELATORIO.FormulaFields.GetItemByName("Cidade").Text = Chr(34) + XCidade + Chr(34)
    XGR_RELATORIO.FormulaFields.GetItemByName("Nominal").Text = Chr(34) + XNominal + Chr(34)
    XGR_RELATORIO.FormulaFields.GetItemByName("Empresa").Text = Chr(34) + PEmpresa + Chr(34)
    FrmVisRelatoriosEXT.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XLA_SUBRELATORIO, "ConsDistribCopiaCheque"
    FrmVisRelatoriosEXT.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XLO_SUBRELATORIO2, "DescAcresc"
    FrmVisRelatoriosEXT.SubVerRelatorio XGR_RELATORIO, ""
    TelaDuplicatas.MousePointer = vbDefault
  End If
  
  'Existem cheques consolidados
  If ChCons Then
    SelecaoCons = SelecaoCons + " and (ConsFINCopiaConsolidado.coco_cd_codigo = " & CboCCorrente2.BoundText & ")"
    
    'Abre os recordsets e verifica se eles possuem dados
    
'*** Verifica qual empresa está executando o sistema para fazer a consulta de acordo com o relatório utilizado por ela.
'** Início - Larissa Moura 21/05/08
    If ((UCase(XLO_EMPRESA!empr_tx_cliente) = "COSTA ANDRADE") Or (UCase(XLO_EMPRESA!empr_tx_cliente) = "RJ") Or (UCase(XLO_EMPRESA!empr_tx_cliente) = "EBISA")) Then
        Sql = "SELECT DISTINCT" & _
                " nofi_tx_historico ," & _
                " nofi_nr_documento, " & _
                " dist_vl_valor, " & _
                " descontoRateado, " & _
                " AcrescimoRateado, " & _
                " dupl_tx_docpagrec, " & _
                " coco_cd_codigo, " & _
                " dupl_vl_valor, " & _
                " acrescimo, " & _
                " desconto, " & _
                " nofi_nr_fatura, " & _
                " dupl_dt_pagRec, " & _
                " focl_tx_razaosocial, " & _
                " coco_cd_agencia, " & _
                " coco_nm_agencia, " & _
                "banc_tx_descricao " & _
                "From ConsFINCopiaConsolidado Where" & SelecaoCons & _
                "AND dist_vl_valor = " & _
                "(select MAX(dist_vl_valor) from ConsFINCopiaConsolidado cf " & _
                "Where cf.dupl_cd_duplicata = ConsFINCopiaConsolidado.dupl_cd_duplicata )"
    Else
        Sql = "SELECT * " & _
            "  From ConsFINCopiaConsolidado Where" & SelecaoCons
        
    End If
'** Fim - Larissa Moura 21/05/08


    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute(Sql)
    If XGO_RSRELATORIO.EOF Then
      MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
      Me.MousePointer = vbDefault
      Exit Sub
    End If
    
    Set XLO_SUBRELATORIO2 = ConexaoRelatorio.Execute("SELECT * FROM ConsFINCChequeDescAcresc WHERE empr_cd_empresa =" & PCodEmpresa)
    
    'Abre o relatório
    FrmVisRelatorios_2.Caption = "Cópia de Cheque Consolidado"
    
    
'*** Verifica qual empresa está executando o sistema para selecionar o relatório que ela utiliza
'** Início - Larissa Moura 21/05/08
    Select Case UCase(XLO_EMPRESA!empr_tx_cliente)
        Case "QUALIDADOS"
            Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\CopiaChequeConsAPAE.rpt")
        Case "FRANISA"
            Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\CopiaChequeConsAPAE.rpt")
        Case "APAE"
            Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\CopiaChequeConsAPAE.rpt")
        Case "ALIANCA"
            Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\CopiaChequeConsAPAE.rpt")
        Case "PLENA"
            Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\CopiaChequeConsAPAE.rpt")
        Case Else 'Costa Andrade, RJ Construções, Ebisa e Temae
            Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\CopiaChequeCons.rpt")
    End Select
'** Fim - Larissa Moura 21/05/08
   
            
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
    
    'Imprime o relatório
    XGR_RELATORIO.FormulaFields.GetItemByName("Cidade").Text = Chr(34) + XCidade + Chr(34)
    XGR_RELATORIO.FormulaFields.GetItemByName("Empresa").Text = Chr(34) + PEmpresa + Chr(34)
    'FrmVisRelatorios_2.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XLO_SUBRELATORIO2, "DescAcresc"
    'FrmVisRelatorios_2.SubVerRelatorio XGR_RELATORIO, ""
    
    'Cássio Medeiros 17/12/2013
    'Essa visualização de relatório permiti selecionar as impressoras disponíveis antes da impressão
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XLO_SUBRELATORIO2, "DescAcresc"
    FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
    TelaDuplicatas.MousePointer = vbDefault
  End If
  
  TelaDuplicatas.MousePointer = vbDefault
  CmdRetornar2_Click
End Sub

Private Sub CmdImprimir_Click()
    Call subImprimeListagemGRID(2, TDBGrid1.PrintInfo, "Listagem de Duplicatas")
End Sub

Private Sub CmdOrdem_Click()
    subDesabilitaBotoes
    subTelaValoresGlobais "P"
    TelaOrdem.Show 1
    subTelaValoresGlobais "G"
'    SQLAtivo = Adodc1.RecordSource
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    subHabilitaBotoes
    lblLocalizar.Caption = "Localizar " & xLocalDescricao & ": "
    cboLocalizar.ListField = xLocalCampo
    cboLocalizar.Text = ""
    TDBGrid1.SetFocus
End Sub

Private Sub CmdPagamento_Click()
  Dim X As Integer
  
  If Not FunVerifica_Permissao("TelaDuplicatas", "1") Then Exit Sub

  'Verificações
  If FunTabelaVazia(Adodc1) Then
    MsgBox "Não pode ser efetivado nenhum pagamento, pois a tabela está vazia", vbCritical, "ATENÇÃO!"
    Exit Sub
  End If
  If Not IsNull(Adodc1.Recordset.Fields("dupl_dt_pagrec")) And Adodc1.Recordset.Fields("dupl_dt_pagrec") <> "" Then
    MsgBox "O pagamento desta duplicata já foi realizado.", vbInformation, "ATENÇÃO!"
    Exit Sub
  End If
  
  'Não permitir efetivação para duplicatas canceladas. 22/10/09 - Patrícia
  If Adodc1.Recordset.Fields("dupl_tx_status") = "X" Then
    MsgBox "Não é possível efetivar o pagamento de uma duplicata cancelada!", vbInformation, "ATENÇÃO!"
    Exit Sub
  End If

  'Não permitir efetivação para registros de Estorno Financeiro
  If Left(Adodc1.Recordset.Fields("Numdocumento"), 1) = "E" Then
      MsgBox "Esta opção não é permitida para registro de Estorno Financeiro!", vbCritical, "ATENÇÃO!"
      Exit Sub
  End If
  'Os registros frutos de um estorno não devem atender a nenhuma solicitação
  If Adodc1.Recordset.Fields("dupl_tx_Estorno") = "S" Then
    MsgBox "Não é possível efetivar um pagamento para um estorno!", vbCritical, "ATENÇÃO!"
    Exit Sub
  End If

  Chave = Adodc1.Recordset.Fields("dupl_cd_duplicata")
  PanPagamento.Left = (TDBGrid1.Width - PanPagamento.Width) / 2
  PanPagamento.Top = (TDBGrid1.Height - PanPagamento.Height) / 2
  If IsNull(Adodc1.Recordset.Fields("Entrada")) Then
    PanPagamento.Caption = "Efetivar Pagamento"
    Label3.Caption = "Pagamento:"
  Else
    PanPagamento.Caption = "Efetivar Recebimento"
    Label3.Caption = "Recebimento:"
  End If
  PanPagamento.Visible = True
  PanPagamento.Enabled = True
  subDesabilitaBotoes
  TDBGrid1.Enabled = False
  For X = 0 To CboTipoPag.ListCount - 1
    If CboTipoPag.ItemData(X) = CInt(Adodc1.Recordset.Fields("dupl_tx_tipopagrec")) Then
      CboTipoPag.ListIndex = X
    End If
  Next
  
  DtpDtPagto.Value = Format$(Now, "DD/MM/YY")
  TxtNumDoc.Text = ""
  'CboCCorrente.BoundText = Adodc1.Recordset.Fields("coco_cd_codigo")
  'Cássio Medeiros 07/10/2014
  CboCCorrente.BoundText = IIf(Not IsNull(Adodc1.Recordset.Fields("coco_cd_codigo")), Adodc1.Recordset.Fields("coco_cd_codigo"), "")
  DtpDtPagto.SetFocus
  CboFormaPagamento.BoundText = IIf(Not IsNull(Adodc1.Recordset.Fields("fopa_cd_codigo")), Adodc1.Recordset.Fields("fopa_cd_codigo"), "")
  
End Sub

Private Sub CmdProcPagto_Click()
  Dim FiltroCCusto As String
  Dim ResNF As ADODB.Recordset
  Dim NumFat As Long
  Dim CodDupl As Long
  Dim XL0_RSSUBRELATORIO As ADODB.Recordset
  Dim XL0_RSSUBRELATORIO2 As ADODB.Recordset
  Dim XLT_EXISTE As String
  Dim XLT_DISTRIBUIDO As String
  
  TelaDuplicatas.MousePointer = vbHourglass
  
  'Verificações
  If Adodc1.Recordset.Fields("dupl_tx_debcred") = "C" Then
    MsgBox "Não existe Processo de Pagamento para um Crédito!", vbCritical, "ATENÇÃO"
    TelaDuplicatas.MousePointer = vbDefault
    Exit Sub
  End If
  
  'Não permitir efetivação para transferências de custo:
  If TDBGrid1.Columns("Tipo Doc.") = "TRANSFERÊNCIA DE CUSTO" Then
      MsgBox "Esta opção não é permitida para um registro de transferência de custo!", vbCritical, "ATENÇÃO!"
      Exit Sub
  End If
  
  XGT_SELECAO = ""
  XGT_CONJUNCAO = ""
  NumFat = Adodc1.Recordset.Fields("dupl_nr_fatura")
  CodDupl = Adodc1.Recordset.Fields("dupl_cd_duplicata")
  FiltroCCusto = ""
  
  'Verifica se Foi distribuído
  SubQOpenRecordset ResNF, "SELECT nofi_tx_distribuir FROM NotasFiscais WHERE nofi_nr_Fatura = " & NumFat & "", Estatico
  If Not ResNF.EOF Then
    If ResNF!nofi_tx_distribuir = "S" Then
        XLT_DISTRIBUIDO = "S"
    Else
        XLT_DISTRIBUIDO = "N"
    End If
  End If
  ResNF.Close
  FiltroCCusto = ""
        
  'Verifica a quantidade de parcelas
  SubQOpenRecordset ResNF, "SELECT dupl_cd_duplicata FROM Duplicatas where dupl_nr_fatura =" & NumFat, "Estatico"
  ResNF.MoveFirst
  Dim Cont As String
  Dim Parcela As String
  Cont = "0"
  While Not ResNF.EOF
    Cont = CStr(CInt(Cont) + 1)
    If ResNF!dupl_cd_Duplicata = CodDupl Then
      Parcela = Cont
    End If
    ResNF.MoveNext
  Wend
  ResNF.Close
        
  XGT_SELECAO = "(ConsFINFluxoCaixa.dupl_cd_duplicata = " & CodDupl & ")" 'AND ConsFINFluxoCaixa.dupl_tx_status = 'C' "
  
  'Abre os recordsets e verifica se eles possuem dados
  Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINFluxoCaixa WHERE " & XGT_SELECAO)
  Set XL0_RSSUBRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDistribCopiaCheque WHERE nofi_nr_fatura=" & Adodc1.Recordset.Fields("dupl_nr_Fatura") & "")
  Set XL0_RSSUBRELATORIO2 = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDescAcresc")
  If XGO_RSRELATORIO.EOF Then
    MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
    Me.MousePointer = vbDefault
    Exit Sub
  End If
  
  'Set XGO_RSCABECALHO = Conexao.Execute("SELECT * FROM Empresas")
  
  'Abre o Relatório
  FrmVisRelatorios.Caption = "Processo de Pagamento"
  If (UCase(PCliente) = "APAE") Then
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\ProcessoPagtoUnicoAPAE.rpt")
  Else
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\ProcessoPagtoUnico.rpt")
  End If
  
  'Passa o recordset para o relatório
  XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO, 3, 1
  
  'Atualiza a tabela de Duplicatas
  SubQOpenRecordset ResNF, "SELECT dupl_dt_ProcPagamento FROM Duplicatas where dupl_cd_Duplicata =" & Adodc1.Recordset.Fields("dupl_cd_Duplicata") & "", Dinamico
  If IsNull(ResNF!dupl_dt_ProcPagamento) Then
      XLT_EXISTE = "N" 'O processo foi impresso pela 1ª vez
      ResNF!dupl_dt_ProcPagamento = Format(Now, "DD/MM/YY")
      ResNF.Update
  Else
      XLT_EXISTE = "S" 'O processo já foi impresso anteriormente
  End If
  ResNF.Close
  
  'Ativa os Relatórios
  With XGR_RELATORIO.FormulaFields
    .GetItemByName("CCusto").Text = Chr(34) + FiltroCCusto + Chr(34)
    .GetItemByName("QuantParcela").Text = Chr(34) + Cont + Chr(34)
    .GetItemByName("Parcela").Text = Chr(34) + Parcela + Chr(34)
    .GetItemByName("Empresa").Text = Chr(34) + PEmpresa + Chr(34)
    .GetItemByName("JaFoiImpresso").Text = Chr(34) + XLT_EXISTE + Chr(34)
    .GetItemByName("Distribuido").Text = Chr(34) + XLT_DISTRIBUIDO + Chr(34)
    If (UCase(PCliente) = "APAE") Then
        .GetItemByName("Emitente").Text = Chr(34) + XGT_USUARIO + Chr(34)
        .GetItemByName("Cargo").Text = Chr(34) + XGT_CARGO + Chr(34)
    End If
  End With
  'FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XL0_RSSUBRELATORIO, "ConsFINDistribCopiaCheque"
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XL0_RSSUBRELATORIO2, "ConsFINDescAcresc"
  FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
  
  '///////Impressão do Anexo do Processo (p/notas distribuidas)//////
  If XLT_DISTRIBUIDO = "S" Then
    Dim XLO_RSRELATORIOANEXO As New ADODB.Recordset
    Dim XLR_RELATORIOANEXO As New CRAXDRT.Report
  
    Set XLO_RSRELATORIOANEXO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDistribCopiaCheque WHERE nofi_nr_fatura=" & NumFat & "")
      
    'Abre o relatório
    FrmVisRelatorios_2.Caption = "Anexo do Processo de Pagamento"
    Set XLR_RELATORIOANEXO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\AnexoProcPagto.rpt")
  
    'Passa o recordset para o relatório
    XLR_RELATORIOANEXO.Database.SetDataSource XLO_RSRELATORIOANEXO, 3, 1
  
    '********** Ativa os Relatórios
    With XLR_RELATORIOANEXO.FormulaFields
        .GetItemByName("NumProcesso").Text = Chr(34) + CStr(NumFat) + Chr(34)
    End With
    FrmVisRelatorios_2.SubVerRelatorio XLR_RELATORIOANEXO, ""
  End If
  '//////////////////////////////////////////////////////////////////
  
  TelaDuplicatas.MousePointer = vbDefault
End Sub

Private Sub CmdRecibo_Click()
    Dim ResTipoDoc As ADODB.Recordset
    Dim XLT_SQL As String
    Dim XLT_HISTORICO As String
    Dim XLO_RSEMPRESA As ADODB.Recordset ' Declaração em 24/09/08 - Patrícia
    
   
    XFT_TEXTO = TxtTextoRecibo.Text
    Modelo = 0
    optModelo1.Visible = False
    optModelo2.Visible = False
    optModelo3.Visible = False
    txtEmissor.Visible = False
    LbLEmissor.Visible = False
    
    If Not IsNull(Adodc1.Recordset.Fields("dupl_dt_PagRec")) Then
        DtpDataCartaTransferencia.Value = Adodc1.Recordset.Fields("dupl_dt_PagRec") '29/09/08 - Patrícia
    End If
    
    PanDetalhesCartaTransferencia.Height = 2100
    PanDetalhesCartaTransferencia.Width = 4530
    
    PanTextoRecibo.Left = (TDBGrid1.Width - PanTextoRecibo.Width) / 2
    PanTextoRecibo.Top = (TDBGrid1.Height - PanTextoRecibo.Height) / 2
    
    PanDetalhesCartaTransferencia.Left = (TDBGrid1.Width - PanDetalhesCartaTransferencia.Width) / 2
    PanDetalhesCartaTransferencia.Top = (TDBGrid1.Height - PanDetalhesCartaTransferencia.Height) / 2
    
    Label11.Top = 720
    txtNumeroRecibo.Top = 720
    LbLEmissor.Visible = False
    Label10.Visible = False
    DtpDataCartaTransferencia.Visible = False
    cmdConfirmarCartaTransferencia.Top = 1440
    cmdConfirmarCartaTransferencia.Left = 1320
    cmdRetornarCartaTransfencia.Left = 2760
    cmdRetornarCartaTransfencia.Top = 1440

    
    'Não permitir efetivação para registros de Estorno Financeiro
    If Left(Adodc1.Recordset.Fields("Numdocumento"), 1) = "E" Then
        MsgBox "Esta opção não  é permitida para registro de Estorno Financeiro!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    
    If Adodc1.Recordset.Fields("dupl_tx_debCred") = "D" Then
        SubQOpenRecordset ResTipoDoc, "SELECT nofi_tx_Historico FROM NotasFiscais WHERE nofi_nr_Fatura = " & Adodc1.Recordset.Fields("dupl_nr_Fatura"), Estatico
        XLT_HISTORICO = ResTipoDoc!nofi_tx_historico
        ResTipoDoc.Close
    
        If Adodc1.Recordset.Fields("dupl_tx_tipopagrec") <> "1" Then
            TxtTextoRecibo.Text = "RECEBEMOS DE " + PEmpresa + " A IMPORTÂNCIA DE R$ " + TDBGrid1.Columns("A Pagar") + " REFERENTE A " + XLT_HISTORICO
        Else
            TxtTextoRecibo.Text = "RECEBEMOS DE " + PEmpresa + " A IMPORTÂNCIA DE R$ " + TDBGrid1.Columns("A Pagar") + " REFERENTE A " + XLT_HISTORICO & "" & _
            ". PAGAMENTO POR CHEQUE Nº " + Adodc1.Recordset.Fields("dupl_tx_DocPagRec") + ", " + Adodc1.Recordset.Fields("BancoEmpresa") + "."
        End If
        
        TxtTextoRecibo.Text = FunTrataString(TxtTextoRecibo.Text)
        
        PanTextoRecibo.Visible = True
        
        Exit Sub
    Else
        SubQOpenRecordset XLO_RSEMPRESA, "SELECT empr_tx_cliente FROM Empresas WHERE empr_cd_empresa = " & CStr(PCodEmpresa), "Estatico"
        If Not (XLO_RSEMPRESA.EOF) Then ' If inserido em 24/09/08 - Patrícia. Inclusão de relatórios solicitados pelo cliente Instituto Aliança
            If Trim(XLO_RSEMPRESA!empr_tx_cliente) = "ALIANCA" Then
                        'Gerar o nº de controle e emitir o recibo
                    PanDetalhesCartaTransferencia.Visible = True
            Else
                    XLT_SQL = "SELECT * FROM Duplicatas WHERE dupl_cd_duplicata = " & Adodc1.Recordset.Fields("dupl_cd_duplicata")
                    SubQOpenRecordset ResTipoDoc, XLT_SQL, Dinamico
                    If IsNull(ResTipoDoc!dupl_tx_controle) Then
                        'Emitir o recibo e gera o nº de controle
                        If MsgBox("Confirma emissão do Recibo?", vbQuestion + vbOKCancel, "ATENÇÃO") = vbOK Then
                            If Adodc1.Recordset.Fields("dupl_tx_DebCred") = "D" Then  'D ou C
                                ResTipoDoc!dupl_tx_controle = CStr(FunPegaGuardaUltimo("C"))
                            Else
                                ResTipoDoc!dupl_tx_controle = CStr(FunPegaGuardaUltimo("R"))
                            End If
                            subImprimeRecibo (ResTipoDoc!dupl_tx_controle)
                        Else
                            Exit Sub
                        End If
                    Else
                        If MsgBox("O recibo já foi emitido. Deseja reemitir ?", vbYesNo + vbCritical, "ATENÇÃO") = vbYes Then
                            'Reemitir sem gerar um novo número de controle
                            subImprimeRecibo (ResTipoDoc!dupl_tx_controle)
                        Else
                            Exit Sub
                        End If
                    End If
                    ResTipoDoc.Update
                    ResTipoDoc.Close
            End If
        End If
        XLO_RSEMPRESA.Close

    End If
End Sub

Private Sub CmdRetorna_Click()
    PanCheque.Left = 9375
    PanCheque.Enabled = False
    PanCheque.Visible = False
    
    TDBGrid1.Enabled = True
    subHabilitaBotoes
    
    'Adodc1.Refresh TestesubrecarregadadosNV
'    SQLAtivo = Adodc1.RecordSource
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    TDBGrid1.Bookmark = xChave
End Sub

Private Sub CmdRetornaEstorno_Click()
    PanEstorno.Visible = False
    PanEstorno.Enabled = False
    subHabilitaBotoes
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
End Sub

Private Sub CmdRetornar_Click()
    PanPagamento.Visible = False
    PanPagamento.Enabled = False
    subHabilitaBotoes
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
End Sub

Private Sub CmdRetornar2_Click()
    TDBGrid1.Enabled = True
    PanCopia.Visible = False
    PanCopia.Enabled = False
    subHabilitaBotoes
    TxtNroCheque2.Text = ""
    TxtNroCheque3.Text = ""
    CboCCorrente2.Text = ""
    CmdSair.Enabled = True
End Sub

Private Sub cmdRetornarCancelamento_Click()
    
    Me.panCancelar.Visible = False
    
End Sub

Private Sub cmdRetornarCartaTransfencia_Click()
    PanDetalhesCartaTransferencia.Visible = False
End Sub

Private Sub CmdRetornarParcial_Click()
    PanBaixaParcial.Visible = False
    PanBaixaParcial.Enabled = False
    subHabilitaBotoes
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
End Sub

Private Sub CmdSair_Click()
    MdiPrincipal.BarraStatus.Panels(3).Text = ""
    MdiPrincipal.BarraStatus.Panels(1).Text = ""
    MdiPrincipal.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub CmdSairDuplicata_Click()
    PanDuplicata.Visible = False
    PanDuplicata.Enabled = False
    subHabilitaBotoes
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
End Sub

Private Sub CmdAltParcela_Click()
    Dim XLI_X As Integer

    'Não permitir efetivação para registros de Estorno Financeiro
    If Left(Adodc1.Recordset.Fields("Numdocumento"), 1) = "E" Then
        MsgBox "Esta opção não é permitida para registro de Estorno Financeiro!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    
    If Adodc1.Recordset.Fields("Status") = "Realizado" Then
        MsgBox "Não é possível alterar os dados de uma parcela já realizada.", vbCritical, "ATENÇÃO"
        Exit Sub
    End If
    
    If Adodc1.Recordset.Fields("TipoDocDesc") = "GLOSA" Or Adodc1.Recordset.Fields("TipoDocDesc") = "NOTA FISCAL DE DEVOLUÇÃO" Then
        If EntSaida = "S" Then
            MsgBox "Não é possível alterar esse tipo de documento.", vbCritical, "ATENÇÃO"
        Else
            MsgBox "Não é possível alterar esse tipo de documento.", vbCritical, "ATENÇÃO"
        End If
        Exit Sub
    End If
    
    'Não permitir alteração de transferências de custo:
    If TDBGrid1.Columns("Tipo Doc.") = "TRANSFERÊNCIA DE CUSTO" Then
        MsgBox "Esta opção não é permitida para um registro de transferência de custo!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    
    xChave = Adodc1.Recordset.Bookmark
    
    'Os registros frutos de um estorno não devem atender a nenhuma solicitação
    If Adodc1.Recordset.Fields("dupl_tx_Estorno") = "S" Then
        MsgBox "Não é permitido alterar um estorno!", vbCritical, "ATENÇÃO!"
        Exit Sub
    End If
    
    If IsNull(Adodc1.Recordset.Fields("focl_tx_fantasia")) Then
    'Se não existe fornecedor o registro é uma transferencia bancária ou recolhimento
        If Adodc1.Recordset.Fields("TipoDocDesc") = "TRANSFERÊNCIA BANCÁRIA" Then
            MsgBox "Não é permitido alterar uma Transferência Bancária.", vbCritical, "ATENÇÃO"
            Exit Sub
        ElseIf Adodc1.Recordset.Fields("TipoDocDesc") = "TRANSFERÊNCIA DE CUSTO" Then
            MsgBox "Não é permitido alterar uma Transferência de Custo.", vbCritical, "ATENÇÃO"
            Exit Sub
        ElseIf Adodc1.Recordset.Fields("TipoDocDesc") = "DAM" _
            Or Adodc1.Recordset.Fields("TipoDocDesc") = "DARF" _
            Or Adodc1.Recordset.Fields("TipoDocDesc") = "INSS" _
            Or Adodc1.Recordset.Fields("TipoDocDesc") = "OUTR" Then
            MsgBox "Recolhimentos não podem ser alterados. Exclua e efetue novo recolhimento.", vbCritical, "ATENÇÃO"
            Exit Sub
        End If
    End If
    
    PanDuplicata.Enabled = True
    PanDuplicata.Left = (TDBGrid1.Width - PanDuplicata.Width) / 2
    PanDuplicata.Top = (TDBGrid1.Height - PanDuplicata.Height) / 2
    DtpVencimento.Value = Adodc1.Recordset.Fields("dupl_dt_vencimento")
    For XLI_X = 0 To CboTipoPagto2.ListCount - 1
        If CboTipoPagto2.ItemData(XLI_X) = CInt(Adodc1.Recordset.Fields("dupl_tx_tipopagrec")) Then
            CboTipoPagto2.ListIndex = XLI_X
        End If
    Next
    CboCCorrente3.BoundText = Adodc1.Recordset.Fields("coco_cd_codigo")
    PanDuplicata.Visible = True

End Sub

Private Sub cmdContabiliza_Click()
    Dim ResFat As ADODB.Recordset
    Dim XTipoFatura As String
    Dim ResTipoDoc As ADODB.Recordset
    Dim xsigla As String

    Dim i As Integer
    
    i = 1
    XGB_CONTABILIZA = True
    
    Me.MousePointer = vbHourglass

    Adodc1.Recordset.MoveFirst

 For i = 1 To Adodc1.Recordset.RecordCount
    

    'Não permitir efetivação para registros de Estorno Financeiro
    If Left(Adodc1.Recordset.Fields("Numdocumento"), 1) <> "E" And _
    TDBGrid1.Columns("Tipo Doc.") <> "TRANSFERÊNCIA DE CUSTO" And _
    IsNull(Adodc1.Recordset.Fields("dupl_tx_Estorno")) Then
    
            XFormulario = "TelaDuplicatas"
            NFatura = Adodc1.Recordset.Fields("dupl_nr_fatura")
            xChave = Adodc1.Recordset.Bookmark
            ChaveDp = Adodc1.Recordset.Fields("dupl_cd_duplicata")
        
        
            SubQOpenRecordset ResFat, "SELECT nofi_tx_controle,nofi_cd_notafiscal,nofi_tx_tipofatura FROM NotasFiscais where nofi_nr_fatura =" & NFatura, "Estatico"
            If Not (ResFat.BOF And ResFat.EOF) Then
                If Not IsNull(ResFat!nofi_tx_controle) Then
                    XControleTransf = ResFat!nofi_tx_controle
                End If
                XTipoFatura = ResFat!nofi_tx_tipofatura
                If XTipoFatura <> "3" Then
                    Chave = ResFat!nofi_cd_notafiscal
                End If
            End If
            ResFat.Close
        
            If IsNull(Adodc1.Recordset.Fields("focl_tx_fantasia")) Then
            'Se não existe fornecedor o registro é uma transferencia bancária ou recolhimento
                SubQOpenRecordset ResTipoDoc, "SELECT TipoDocumentos.tido_tx_sigla FROM TipoDocumentos,NotasFiscais WHERE NotasFiscais.nofi_nr_fatura = " & NFatura & " and notasfiscais.tido_cd_tipodoc = TipoDocumentos.tido_cd_tipodoc ", "Estatico"
                If Not ResTipoDoc.EOF Then
                    xsigla = FunNulo(ResTipoDoc!tido_tx_sigla)
                Else
                    MsgBox "Não exite duplicata com este nº de PP. Atualize a tela!", vbCritical, "ATENÇÃO"
                    Exit Sub
                End If
                ResTipoDoc.Close
        
                If xsigla = "TB" Then
                    GoTo NaoExecuta
                ElseIf (xsigla = "TBCP" Or xsigla = "TBCR") Then
                    GoTo NaoExecuta
                ElseIf xsigla = "TRCC" Then
                    GoTo NaoExecuta
                ElseIf xsigla = "DAM" Or xsigla = "DARF" Or xsigla = "INSS" Or xsigla = "OUTR" Then
                    GoTo NaoExecuta
                End If
            End If
        
            XInserir = "A"
            If Not IsNull(Adodc1.Recordset.Fields("Saida")) Then
                EntSaida = "S"
            ElseIf Not IsNull(Adodc1.Recordset.Fields("Entrada")) Then
                EntSaida = "E"
            End If
        
            On Error Resume Next
            If XTipoFatura = "3" Then
                GoTo NaoExecuta
            Else
                FrmNotasFiscais.Show 1
            End If
        
            If FrmNotasFiscais.CARREGAGRID = True Then
                subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
                TDBGrid1.Refresh
                If Not Adodc1.Recordset.EOF Then
                    If Adodc1.Recordset.RecordCount > xChave Then
                        Adodc1.Recordset.Bookmark = xChave
                    End If
                End If
        
            End If
    
    
    End If
    
NaoExecuta:
     On Error Resume Next
    Adodc1.Recordset.MoveNext

Next i

    XGB_CONTABILIZA = False
    
    Me.MousePointer = vbDefault
    
End Sub

Private Sub Command2_Click()
    Me.panCancelar.Visible = False
End Sub

Private Sub CmdSairReciboAlianca_Click()
    PanDetalhesCartaTransferencia.Visible = False
    txtNumeroRecibo.Text = ""
End Sub

Private Sub dtpVencimento_LostFocus()
  Dim XLO_NF As ADODB.Recordset
  
  SubQOpenRecordset XLO_NF, "SELECT nofi_dt_emissao FROM NotasFiscais " & _
            " WHERE nofi_nr_fatura = " & Adodc1.Recordset.Fields("dupl_nr_fatura"), Estatico
  
  If Not XLO_NF.EOF Then
    If DateDiff("d", XLO_NF!nofi_dt_emissao, DtpVencimento.Value) < 0 Then
        MsgBox "A data de vencimento é anterior a data de emissão da Nota !", vbCritical, "ATENÇÃO"
        If DtpVencimento.Enabled = True Then DtpVencimento.SetFocus
        Exit Sub
    End If
  End If
  XLO_NF.Close
End Sub




Private Sub optModelo1_Click()
    txtEmissor.Enabled = True
    optModelo2.Value = False
    optModelo3.Value = False
End Sub

Private Sub optModelo2_Click()
    txtEmissor.Enabled = False
    optModelo1.Value = False
    optModelo3.Value = False

End Sub


Private Sub optModelo3_Click()
    txtEmissor.Enabled = False
    optModelo1.Value = False
    optModelo2.Value = False

End Sub

Private Sub TDBGrid1_Click()
   If Not Adodc1.Recordset.EOF Then
        If Adodc1.Recordset.Fields("dupl_tx_status") <> "R" Then
             CmdEstorno.Enabled = False
             CmdPagamento.Enabled = True
             CmdCancelar.Enabled = False
             If Adodc1.Recordset.Fields("ExpBanco") = "Sim" Then
                 CmdCanExpBanco.Enabled = True
             Else
                 CmdCanExpBanco.Enabled = False
             End If
         Else
             CmdCanExpBanco.Enabled = False
             CmdPagamento.Enabled = False
             CmdCancelar.Enabled = True
         End If
         
         If Adodc1.Recordset.Fields("TipoRecibo") = "1" Then
             CmdRecibo.Enabled = True
             cmdCartaTransferencia.Enabled = True
         Else
             CmdRecibo.Enabled = False
             cmdCartaTransferencia.Enabled = False
         End If
         
         
         If Adodc1.Recordset.Fields("dupl_tx_Estorno") = "S" Then
             CmdCancelar.Enabled = False
             CmdPagamento.Enabled = False
         End If
         
         If TDBGrid1.Columns("Tipo Pagto./Rec.").Text = "Dinheiro" Then
             CmdEstorno.Enabled = False
         End If
                
         If Adodc1.Recordset.Fields("dupl_tx_debcred") <> "D" Then
             CmdProcPagto.Enabled = False
         Else
            CmdProcPagto.Enabled = True
         End If
         
         If TDBGrid1.Columns("Estoque").Text = "Pendente" Then
             CmdPagamento.Enabled = False
             CmdProcPagto.Enabled = False
         End If
    End If

End Sub

Private Sub TDBGrid1_DblClick()
  
  CmdAlterar_Click
  
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    If TDBGrid1.Columns("A Pagar").CellText(Bookmark) <> "" Then
        RowStyle.ForeColor = &HFF&
    ElseIf TDBGrid1.Columns("A Receber").CellText(Bookmark) <> "" Then
        RowStyle.ForeColor = &H0&
    End If
End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cboLocalizar.Text = ""
    If Button = 2 Then  'Verifica se o botão da direita foi pressionado
        Set Formulario = TelaDuplicatas
        subTelaValoresGlobais "P"
        FunExecutaFiltro Adodc1, Formulario, X
        subTelaValoresGlobais "G"
        Filtro = True
    End If
    
    subHabilitaBotoes
    
    If Filtro = True Then
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
        Filtro = False 'Para o refresh
    End If
        
    If FunTabelaVazia(Adodc1) = True Then
        CmdAlterar.Enabled = False
        CmdPagamento.Enabled = False
        CmdDescAcresc.Enabled = False
        CmdCancelar.Enabled = False
        CmdImprimir.Enabled = False
        CmdOrdem.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
         Set Formulario = TelaDuplicatas
     If Not PrimeiraVez Then
          subTelaValoresGlobais "P"
     Else
          Adodc1.Recordset.Find "dupl_dt_vencimento >=" & Format$(Now, "DD/MM/YY")
          If Adodc1.Recordset.EOF And Not FunTabelaVazia(Adodc1) Then
             Adodc1.Recordset.MoveLast
          End If
     End If
     
     PrimeiraVez = False
     
     subHabilitaBotoes
     
    
    If XGB_PQUALIEST = True Then
        TDBGrid1.Columns("Estoque").Visible = True
    End If

    MdiPrincipal.BarraStatus.Panels(3).Text = CStr(Adodc1.Recordset.RecordCount) + "/" + CStr(Adodc1.Recordset.RecordCount)
End Sub

Private Sub Form_Load()
    Dim XLD_DATAENTRADA     As Date
    Dim XLI_QTDDIASFILTRO   As Integer

    Set Formulario = TelaDuplicatas
    Call subAjustaTelaLoc
    subManutencaoJanelasAtivas "I", "TelaDuplicatas"
    
    cmdContabiliza.Visible = (XGT_USUARIO = "Suporte") Or ((XGT_USUARIO = "Qualidados"))
    
    'Verifica se o número de registros a serem visualizados na tela está definido no arquivo QualiAdmFin.ini. Se não estiver definido no arquivo, procura na tabela de empresas (campo empr_nr_QtdDiasFiltro).
    If Val(Trim(XGT_NUMDIAS)) <> 0 Then
        XLI_QTDDIASFILTRO = Val(Trim(FunNulo(XGT_NUMDIAS)))
        XLD_DATAENTRADA = DateAdd("d", (-Val(XGT_NUMDIAS)), Format(Now, "dd/MM/yyyy"))
    Else
        XLI_QTDDIASFILTRO = XGI_QTDDIASFILTRO
        XLD_DATAENTRADA = DateAdd("d", (-XGI_QTDDIASFILTRO), Format(Now, "dd/MM/yyyy"))
    End If
    
    
    
    If XLI_QTDDIASFILTRO > 0 Then
    'Lucas, 14.01.09 - Alterado para selecionar apenas a quantidade configurada pelo usuário na variável
    'XGI_QTDDIASFILTRO e para ordenar por dupl_dt_vencimento em ordem decrescente para mostrar os ultimos dados
        subCarregaDadosNV Adodc1, "ConsFINDuplicatas", "dupl_dt_vencimento DESC", _
                                "", "SELECT TOP " & XLI_QTDDIASFILTRO & " * FROM ConsFINDuplicatas " & _
                                " WHERE ((status = 'Realizado' " & _
                                " AND dupl_dt_pagrec >= " & FunNuloData(XLD_DATAENTRADA, NomeSgbd) & ") " & _
                                " OR status = 'Compromissado' " & _
                                " OR status = 'Cancelado' " & _
                                " OR status = 'Previsto') " & _
                                " AND empr_cd_empresa =" & PCodEmpresa & _
                                " AND nofi_tx_status <> 'X' " & _
                                "" & IIf(Not FunVerifica_Permissao_SemMsg("TelaNotasFiscais", "12"), "AND (nofi_bl_Sigiloso <> 1 OR nofi_bl_Sigiloso IS NULL) ", "") & ""
    Else
        subCarregaDadosNV Adodc1, "ConsFINDuplicatas", "dupl_dt_vencimento", _
                                "", "SELECT * FROM ConsFINDuplicatas " & _
                                " WHERE (status = 'Realizado' " & _
                                " OR status = 'Cancelado' " & _
                                " OR status = 'Compromissado' " & _
                                " OR status = 'Previsto') " & _
                                " AND empr_cd_empresa =" & PCodEmpresa & _
                                " AND nofi_tx_status <> 'X' " & _
                                "" & IIf(Not FunVerifica_Permissao_SemMsg("TelaNotasFiscais", "12"), "AND (nofi_bl_Sigiloso <> 1 OR nofi_bl_Sigiloso IS NULL) ", "") & ""
    End If
    
    filtra_especial = True
    Filtro = False
    subHabilitaBotoes
    
    'Alteração busca CGCRaiz Samuel - 27/09/18
    If Mid(PEmpresa, 1, 10) = "QUALIDADOS" Or Mid(PEmpresa, 1, 8) = "QUALINFO" Then
        subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_CGCRaiz = '" & Mid(XGT_CGC, 1, 10) & "' AND coco_tx_status = 'A' ORDER BY coco_tx_Descricao", Estatico
    Else
        subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo where empr_cd_empresa=" & PCodEmpresa & " AND coco_tx_status = 'A' ORDER BY coco_tx_Descricao", Estatico
    End If
    subConectarControleDadosNV Datnominal, "SELECT * FROM Bancos ", Estatico
    subConectarControleDadosNV DatTipoDoc, "SELECT * FROM Tipodocumentos WHERE ((tido_tx_sigla IS NULL) or tido_tx_sigla='RA' or tido_tx_sigla='NF' or tido_tx_sigla='RGLO' or tido_tx_sigla='NFCF' or tido_tx_sigla='NFF' or tido_tx_sigla='NFS'  or tido_tx_sigla='NFFS') ORDER BY tido_tx_descricao", Estatico
    
    subTelaValoresGlobais "G"
    PrimeiraVez = True
    If FunTabelaVazia(Adodc1) Then
       CmdAlterar.Enabled = False
       CmdCancelar.Enabled = False
       CmdPagamento.Enabled = False
       CmdDescAcresc.Enabled = False
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
       CmdEstorno.Enabled = False
       CmdCopia.Enabled = False
       CmdProcPagto.Enabled = False
       CmdRecibo.Enabled = False
       cmdCartaTransferencia.Enabled = False
    Else
        If Adodc1.Recordset.EOF Then
            Adodc1.Recordset.MoveLast
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "TelaDuplicatas"
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If Not Adodc1.Recordset.EOF Then
    
        If Adodc1.Recordset.Fields("dupl_tx_status") <> "R" Then
             CmdEstorno.Enabled = False
             CmdPagamento.Enabled = True
             CmdCancelar.Enabled = False
             If Adodc1.Recordset.Fields("ExpBanco") = "Sim" Then
                 CmdCanExpBanco.Enabled = True
             Else
                 CmdCanExpBanco.Enabled = False
             End If
         Else
             CmdCanExpBanco.Enabled = False
             CmdPagamento.Enabled = False
             CmdCancelar.Enabled = True
         End If
         
         If Adodc1.Recordset.Fields("TipoRecibo") = "1" Then
            CmdRecibo.Enabled = True
            cmdCartaTransferencia.Enabled = True
         Else
            CmdRecibo.Enabled = False
            cmdCartaTransferencia.Enabled = False
         End If
         
         If Adodc1.Recordset.Fields("dupl_tx_Estorno") = "S" Then
             CmdCancelar.Enabled = False
             CmdPagamento.Enabled = False
         End If
         
         If TDBGrid1.Columns("Tipo Pagto./Rec.").Text = "Dinheiro" Then
             CmdEstorno.Enabled = False
         Else
             CmdEstorno.Enabled = True
         End If
         
         If Adodc1.Recordset.Fields("dupl_tx_debcred") <> "D" Then
             CmdProcPagto.Enabled = False
         Else
            CmdProcPagto.Enabled = True
         End If
         
         If TDBGrid1.Columns("Estoque").Text = "Pendente" Then
             CmdPagamento.Enabled = False
             CmdProcPagto.Enabled = False
         End If
         
         If Adodc1.Recordset.Fields("TipoDocDesc") = "NOTA PROMISSÓRIA" Then
            CmdBaixaParcial.Enabled = False
            CmdAltParcela.Enabled = False
            CmdDescAcresc.Enabled = False
            CmdCopia.Enabled = False
            CmdPagamento.Enabled = False
            CmdProcPagto.Enabled = False
            CmdEstorno.Enabled = False
            CmdCancelar.Enabled = False
            cmdCancelarDuplicata.Enabled = False
            CmdCanExpBanco.Enabled = False
         Else
            CmdBaixaParcial.Enabled = True
            CmdAltParcela.Enabled = True
            CmdDescAcresc.Enabled = True
            CmdCopia.Enabled = True
            CmdAlterar.Enabled = True
         End If
    End If
    
End Sub
Private Sub TDBGrid1_SelChange(Cancel As Integer)

    Dim XLD_TOTALSELECAO As Variant
    Dim i As Integer

    i = 0
    
    XLD_TOTALSELECAO = 0
       
    If TDBGrid1.SelRange Then
        If TDBGrid1.SelStartCol = TDBGrid1.SelEndCol And (TDBGrid1.SelEndCol = 3 Or TDBGrid1.SelEndCol = 4 Or TDBGrid1.SelEndCol = 5) Then
        
            For i = 0 To TDBGrid1.SelBookmarks.Count - 1 Step 1
            
                If (TDBGrid1.Columns(TDBGrid1.SelEndCol).CellValue(TDBGrid1.SelBookmarks(i))) <> "" Then
                    XLD_TOTALSELECAO = XLD_TOTALSELECAO + (TDBGrid1.Columns(TDBGrid1.SelEndCol).CellValue(TDBGrid1.SelBookmarks(i)))
                End If
            Next
            
            MsgBox "Total selecionado: R$ " & Format(CStr(XLD_TOTALSELECAO), "0.00"), vbInformation, "Seleção"
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
Private Sub TxtCCorrente3_GotFocus()
    Call subSelecionaTXT(TxtCCorrente3)
End Sub
Private Sub TxtCCorrente3_LostFocus()
    If TxtCCorrente3.Text <> "" Then
        CboCCorrente3.BoundText = TxtCCorrente3.Text
        If CboCCorrente3.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            TxtCCorrente3.SetFocus
        End If
    End If
End Sub
Private Sub TxtCodCCorrenteResiduo_GotFocus()
    Call subSelecionaTXT(TxtCodCCorrenteResiduo)
End Sub
Private Sub TxtCodCCorrenteResiduo_LostFocus()
    If TxtCCorrente3.Text <> "" Then
        CboCCorrenteResiduo.BoundText = TxtCodCCorrenteResiduo.Text
        If CboCCorrenteResiduo.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            TxtCodCCorrenteResiduo.SetFocus
        End If
    End If
End Sub
Private Sub TxtCodCCorrenteParcial_GotFocus()
    Call subSelecionaTXT(TxtCodCCorrenteParcial)
End Sub
Private Sub TxtCodCCorrenteParcial_LostFocus()
    If TxtCodCCorrenteParcial.Text <> "" Then
        CboCCorrenteParcial.BoundText = TxtCodCCorrenteParcial.Text
        If CboCCorrenteParcial.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            TxtCodCCorrenteParcial.SetFocus
        End If
    End If
End Sub
Private Sub TxtCodBanco_GotFocus()
    Call subSelecionaTXT(TxtCodBanco)
End Sub


Private Sub TxtCodBanco_LostFocus()
    If TxtCodBanco.Text <> "" Then
        CboNominal.BoundText = TxtCodBanco.Text
        If CboNominal.Text = "" Then
            MsgBox "Não existe Banco com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            TxtCodBanco.SetFocus
        End If
        TxtCodBanco.Text = Format(TxtCodBanco.Text, "0000")
    End If
End Sub


Private Sub TxtNroCheque_LostFocus()
    TxtNroCheque.Text = Format(TxtNroCheque, "0000000")
End Sub

Private Sub TxtNroCheque2_LostFocus()
    TxtNroCheque2.Text = Format(TxtNroCheque2, "0000000")
    TxtNroCheque3.Text = TxtNroCheque2.Text
End Sub

Private Sub TxtNroCheque3_LostFocus()
    TxtNroCheque3.Text = Format(TxtNroCheque3, "0000000")
End Sub

Private Sub TxtNumDoc_LostFocus()
    If CboTipoPag.ListIndex = 0 Then
        TxtNumDoc.Text = Format(TxtNumDoc, "0000000")
    End If
End Sub

Private Sub CmdSairRecibo_Click()
    PanTextoRecibo.Visible = False
End Sub
Private Sub TxtValorParcial_keypress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtValorParcial_LostFocus()
    If IsNumeric(TxtValorParcial.Text) Then
        TxtValorParcial.Text = Format$(TxtValorParcial.Text, "##,##0.00")
        
        If CCur(TxtValorParcial.Text) > Adodc1.Recordset.Fields("Valor") + Adodc1.Recordset.Fields("DA") Then
            MsgBox "O valor da Baixa Parcial não pode exceder o valor Total da Parcela.", vbCritical, "ATENÇÃO"
            TxtValorParcial.SetFocus
            Exit Sub
        End If
    Else
        TxtValorParcial.Text = ""
    End If
End Sub

