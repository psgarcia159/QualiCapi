VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTransferenciaComCusto 
   Caption         =   "Transferência Bancária com Apropriação de Custo"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   LinkTopic       =   "Transferência Bancária com Apropriação de Custo"
   ScaleHeight     =   6105
   ScaleWidth      =   10545
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab sstTab 
      Height          =   5175
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Transferência"
      TabPicture(0)   =   "frmTransferenciaComCusto.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label21"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DatBanco"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "DatPCR"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "DatConta"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "DatContaCorrente"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "CmdLimparHistorico"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "CboHistorico"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "DatHistorico"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "FraBcOrigem"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "DtpDtTransf"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtReferencia"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "TxtDocPagto"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "CboTipoPag"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "TxtVlTransf"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "CboStatus"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "FraBcDestino"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Distribuição"
      TabPicture(1)   =   "frmTransferenciaComCusto.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label13"
      Tab(1).Control(1)=   "LblValorNota"
      Tab(1).Control(2)=   "Label9"
      Tab(1).Control(3)=   "LblValorTotalDistrib"
      Tab(1).Control(4)=   "tdbDistribuicao"
      Tab(1).Control(5)=   "cmdInserirContaCorrenteDestino"
      Tab(1).Control(6)=   "cmdAlterarContaCorrenteDestino"
      Tab(1).Control(7)=   "cmdExcluirContaCorrenteDestino"
      Tab(1).Control(8)=   "pnlDistribuicao"
      Tab(1).ControlCount=   9
      Begin Threed.SSFrame FraBcDestino 
         Height          =   1035
         Left            =   240
         TabIndex        =   42
         Top             =   2520
         Width           =   8670
         _Version        =   65536
         _ExtentX        =   15293
         _ExtentY        =   1826
         _StockProps     =   14
         Caption         =   "Banco Destino"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShadowStyle     =   1
         Begin VB.TextBox TxtCCorrenteDestino 
            Height          =   330
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   10
            Top             =   195
            Width           =   450
         End
         Begin MSMask.MaskEdBox MskContaDestino 
            Height          =   330
            Left            =   1680
            TabIndex        =   12
            Top             =   600
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   7
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo CboPlanoContasDestino 
            Bindings        =   "frmTransferenciaComCusto.frx":0038
            Height          =   330
            Left            =   2640
            TabIndex        =   13
            Top             =   600
            Width           =   5955
            _ExtentX        =   10504
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
         Begin MSDataListLib.DataCombo CboCCorrenteDestino 
            Bindings        =   "frmTransferenciaComCusto.frx":004F
            Height          =   330
            Left            =   2205
            TabIndex        =   11
            Top             =   195
            Width           =   6360
            _ExtentX        =   11218
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
         Begin VB.Label Label11 
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
            Height          =   210
            Left            =   150
            TabIndex        =   44
            Top             =   615
            Width           =   1485
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
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
            Height          =   225
            Left            =   165
            TabIndex        =   43
            Top             =   255
            Width           =   1470
         End
      End
      Begin Threed.SSPanel pnlDistribuicao 
         Height          =   1455
         Left            =   -75000
         TabIndex        =   35
         Top             =   2280
         Visible         =   0   'False
         Width           =   9735
         _Version        =   65536
         _ExtentX        =   17171
         _ExtentY        =   2566
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
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   2
         Alignment       =   0
         Begin VB.TextBox TxtPercDistribuicao 
            Height          =   330
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   46
            Top             =   840
            Width           =   660
         End
         Begin VB.CommandButton cmdRetornarContaCorrenteDestino 
            Caption         =   "Retornar"
            Height          =   330
            Left            =   8280
            TabIndex        =   24
            Top             =   960
            Width           =   1080
         End
         Begin VB.CommandButton cmdConfirmarContaCorrenteDestino 
            Caption         =   "Confirmar"
            Height          =   330
            Left            =   7080
            TabIndex        =   22
            Top             =   960
            Width           =   1080
         End
         Begin VB.TextBox txtValorDistribuicao 
            Height          =   330
            Left            =   3000
            TabIndex        =   21
            Top             =   840
            Width           =   1455
         End
         Begin MSMask.MaskEdBox txtCodigoCentroCustoDestino 
            Height          =   330
            Left            =   1515
            TabIndex        =   19
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   7
            Mask            =   "#####-#"
            PromptChar      =   " "
         End
         Begin Threed.SSCommand cmdLimparCentroCustoDestino 
            Height          =   315
            Left            =   8985
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   240
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "frmTransferenciaComCusto.frx":006E
            Picture         =   "frmTransferenciaComCusto.frx":04C0
         End
         Begin MSDataListLib.DataCombo cboCentroCustoDestino 
            Bindings        =   "frmTransferenciaComCusto.frx":05D2
            DataSource      =   "datCentroCustoDestino"
            Height          =   330
            Left            =   2400
            TabIndex        =   20
            Top             =   240
            Width           =   6570
            _ExtentX        =   11589
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
         Begin VB.Label LblPercentual 
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
            Left            =   360
            TabIndex        =   45
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label14 
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
            Left            =   120
            TabIndex        =   37
            Top             =   300
            Width           =   1350
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2400
            TabIndex        =   36
            Top             =   720
            Width           =   570
         End
      End
      Begin VB.CommandButton cmdExcluirContaCorrenteDestino 
         Caption         =   "Excl&uir"
         Height          =   330
         Left            =   -66600
         TabIndex        =   41
         Top             =   1440
         Width           =   1080
      End
      Begin VB.CommandButton cmdAlterarContaCorrenteDestino 
         Caption         =   "&Al&terar"
         Height          =   330
         Left            =   -66600
         TabIndex        =   40
         Top             =   960
         Width           =   1080
      End
      Begin VB.CommandButton cmdInserirContaCorrenteDestino 
         Caption         =   "Ins&erir"
         Height          =   330
         Left            =   -66600
         TabIndex        =   38
         Top             =   480
         Width           =   1080
      End
      Begin VB.ComboBox CboStatus 
         Height          =   315
         ItemData        =   "frmTransferenciaComCusto.frx":05E7
         Left            =   4320
         List            =   "frmTransferenciaComCusto.frx":05F1
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   2385
      End
      Begin VB.TextBox TxtVlTransf 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7515
         MaxLength       =   50
         TabIndex        =   3
         Top             =   600
         Width           =   1395
      End
      Begin VB.ComboBox CboTipoPag 
         Height          =   315
         ItemData        =   "frmTransferenciaComCusto.frx":060F
         Left            =   2145
         List            =   "frmTransferenciaComCusto.frx":0628
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   990
         Width           =   2955
      End
      Begin VB.TextBox TxtDocPagto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7515
         MaxLength       =   7
         TabIndex        =   5
         Top             =   990
         Width           =   1395
      End
      Begin VB.TextBox txtReferencia 
         Height          =   735
         Left            =   1920
         MaxLength       =   400
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   4200
         Width           =   6855
      End
      Begin MSComCtl2.DTPicker DtpDtTransf 
         Height          =   315
         Left            =   2145
         TabIndex        =   1
         Top             =   600
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         DateIsNull      =   -1  'True
         Format          =   83034113
         CurrentDate     =   37460
      End
      Begin Threed.SSFrame FraBcOrigem 
         Height          =   1020
         Left            =   240
         TabIndex        =   23
         Top             =   1440
         Width           =   8670
         _Version        =   65536
         _ExtentX        =   15293
         _ExtentY        =   1799
         _StockProps     =   14
         Caption         =   "Banco Origem"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShadowStyle     =   1
         Begin VB.TextBox TxtCCorrenteOrigem 
            Height          =   330
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   6
            Top             =   210
            Width           =   450
         End
         Begin MSMask.MaskEdBox MskContaOrigem 
            Height          =   330
            Left            =   1680
            TabIndex        =   8
            Top             =   600
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   7
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo CboPlanoContasOrigem 
            Bindings        =   "frmTransferenciaComCusto.frx":0682
            Height          =   330
            Left            =   2640
            TabIndex        =   9
            Top             =   600
            Width           =   5955
            _ExtentX        =   10504
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
         Begin MSDataListLib.DataCombo CboCCorrenteOrigem 
            Bindings        =   "frmTransferenciaComCusto.frx":0699
            Height          =   330
            Left            =   2205
            TabIndex        =   7
            Top             =   210
            Width           =   6360
            _ExtentX        =   11218
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
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
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
            Height          =   225
            Left            =   300
            TabIndex        =   26
            Top             =   255
            Width           =   1335
         End
         Begin VB.Label Label8 
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
            Height          =   210
            Left            =   150
            TabIndex        =   25
            Top             =   600
            Width           =   1485
         End
      End
      Begin MSAdodcLib.Adodc DatHistorico 
         Height          =   330
         Left            =   9120
         Top             =   2760
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
      Begin MSDataListLib.DataCombo CboHistorico 
         Bindings        =   "frmTransferenciaComCusto.frx":06B8
         Height          =   315
         Left            =   1920
         TabIndex        =   14
         Top             =   3720
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         ListField       =   "hist_tx_descricao"
         BoundColumn     =   "hist_cd_historico"
         Text            =   ""
      End
      Begin Threed.SSCommand CmdLimparHistorico 
         Height          =   330
         Left            =   8520
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   3720
         Width           =   360
         _Version        =   65536
         _ExtentX        =   635
         _ExtentY        =   582
         _StockProps     =   78
         Enabled         =   0   'False
         MouseIcon       =   "frmTransferenciaComCusto.frx":06D3
         Picture         =   "frmTransferenciaComCusto.frx":0B25
      End
      Begin MSAdodcLib.Adodc DatContaCorrente 
         Height          =   330
         Left            =   9120
         Top             =   3840
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
         Left            =   9120
         Top             =   4200
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
      Begin MSAdodcLib.Adodc DatPCR 
         Height          =   330
         Left            =   9120
         Top             =   3120
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
         Caption         =   "DatPCR"
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
      Begin MSAdodcLib.Adodc DatBanco 
         Height          =   330
         Left            =   9120
         Top             =   3480
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
      Begin TrueDBGrid70.TDBGrid tdbDistribuicao 
         Height          =   2895
         Left            =   -74760
         TabIndex        =   39
         Top             =   480
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   5106
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
         Columns(2).Caption=   "Valor"
         Columns(2).DataField=   ""
         Columns(2).NumberFormat=   "Standard"
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
         Splits(0)._ColumnProps(18)=   "Column(2).Width=2566"
         Splits(0)._ColumnProps(19)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(2)._WidthInPix=2487"
         Splits(0)._ColumnProps(21)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(2)._ColStyle=1049362"
         Splits(0)._ColumnProps(23)=   "Column(2).WrapText=1"
         Splits(0)._ColumnProps(24)=   "Column(2).AllowFocus=0"
         Splits(0)._ColumnProps(25)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(26)=   "Column(2)._MinWidth=86573588"
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
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=1,.valignment=2"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44,.alignment=1"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
         _StyleDefs(48)  =   "Named:id=33:Normal"
         _StyleDefs(49)  =   ":id=33,.parent=0"
         _StyleDefs(50)  =   "Named:id=34:Heading"
         _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(52)  =   ":id=34,.wraptext=-1"
         _StyleDefs(53)  =   "Named:id=35:Footing"
         _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=36:Selected"
         _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=37:Caption"
         _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(59)  =   "Named:id=38:HighlightRow"
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
      End
      Begin VB.Label LblValorTotalDistrib 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -68760
         TabIndex        =   50
         Top             =   4560
         Width           =   1485
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   210
         Left            =   -69840
         TabIndex        =   49
         Top             =   4605
         Width           =   1170
      End
      Begin VB.Label LblValorNota 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -72480
         TabIndex        =   48
         Top             =   4560
         Width           =   1485
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor da Transferência:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   210
         Left            =   -74520
         TabIndex        =   47
         Top             =   4605
         Width           =   2010
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Height          =   225
         Left            =   570
         TabIndex        =   34
         Top             =   1035
         Width           =   1560
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Transferência:"
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
         Left            =   360
         TabIndex        =   33
         Top             =   615
         Width           =   1770
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Height          =   225
         Left            =   3585
         TabIndex        =   32
         Top             =   630
         Width           =   720
      End
      Begin VB.Label Label3 
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
         Height          =   225
         Left            =   6810
         TabIndex        =   31
         Top             =   615
         Width           =   675
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Doc. Pagto.:"
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
         Left            =   6360
         TabIndex        =   30
         Top             =   1035
         Width           =   1125
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Referente:"
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
         Left            =   840
         TabIndex        =   29
         Top             =   4215
         Width           =   1005
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
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
         Height          =   225
         Left            =   900
         TabIndex        =   28
         Top             =   3765
         Width           =   885
      End
   End
   Begin Threed.SSCommand CmdGravar 
      Height          =   315
      Left            =   2280
      TabIndex        =   16
      Top             =   5520
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   556
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
   Begin Threed.SSCommand CmdDesistir 
      Height          =   315
      Left            =   4080
      TabIndex        =   17
      Top             =   5520
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   556
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
End
Attribute VB_Name = "frmTransferenciaComCusto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XContaOrigem As String 'Armazena o nº da conta origem
Dim XCCOrigem As String 'Armazena o código da conta contábil origem
Dim XPcrOrigem As String 'Armazena o código do Pcr do banco origem
Dim XContaDestino As String 'Armazena o nº da conta destino
Dim XCCDestino As String 'Armazena o código da conta contábil destino
Dim XPcrDestino As String 'Armazena o código do Pcr do banco destino


Dim XFL_FATURA As Long 'Armazena nº da fatura correspondente à despesa
Dim XFT_INSERIR As String
Dim arrDistribuicao As New XArray

'Variáveis utilizadas na gravação do formulário
Dim Sql As String
Dim XTipoDocComCusto As Integer
Dim XTipoDocSemCusto As Integer
Dim XControle As Long
Dim XFL_CODDUPL As Long     'Armazena o código de pelo 1 das partes da duplicata para utilizar na exp. para contabilidade
Dim XLT_NAOSAIRELFINANCEIRO As String  'Indica se a transferencia sairá ou não nos relatórios financeiros
Dim ValorTotalDistribuicao As Double 'Acumula o total de distribuições até compor o valor total da Nota Fiscal
Dim EntrouDist As Boolean
Dim XLT_CodigoCentroCustoDestinoANT As String
Dim XLT_CentroCustoDestinoANT As String
Dim XLT_ValorDistribuicaoANT As String
Dim Vetor3 As New XArray 'Armazena os dados da distribuição (CRS e valores distribuídos)
Dim PrimeiraAlteracao As Boolean 'Será = True quando o botão Alterar da tela de Notas Fiscais (ou Duplicatas) for pressionado e for a primeira alteração nos registros da tabela de Distribuição(tdbDistribuicao).
Dim PrimeiraInsercao As Boolean 'Será = True quando o botão Alterar da tela de Notas Fiscais (ou Duplicatas) for pressionado e for a primeira alteração nos registros da tabela de Distribuição(tdbDistribuicao).
Dim ValorSomaVetor As Double 'somatório das distribuições armazenadas no grid
Dim ValorSomaVetorANT As Double 'somatório das distribuições armazenadas no grid
Dim XPcr As String  'Armazena o pcr digitado na distribuicao
Dim XDistribuicao As String 'Armazena o tipo de operação (A:alteração, I: inserção)do grid que será executada
Dim XDistribuido As String 'verifica se o valor total foi distribuido

Sub subPreencheCampos()
    Dim Sql As String
    Dim ResTransf As Object
    Dim Status As Integer
    Dim X As Integer

    
'       Sql = "SELECT * FROM ConsFINTransfBancaria WHERE nofi_tx_controle ='" & XControleTransf & "' AND left(nofi_nr_documento, 1) = 'R'  and empr_cd_empresa = " & PCodEmpresa & ""
'       subCarregaVetor Sql, Array("cere_cd_pcr", "cere_cd_estruturado", "plco_cd_conta", "plco_cd_codreduzido", "nofi_tx_historico", "dupl_tx_status", "dupl_tx_TipoPagRec", "dupl_tx_DocPagRec", "dupl_vl_valor"), Vetor3, TDBGrid3
'
    
    SubQOpenRecordset ResTransf, "SELECT * FROM ConsFINTransfBancaria WHERE nofi_tx_controle ='" & XControleTransf & "' and empr_cd_empresa = " & PCodEmpresa & "", "Estatico"
    ResTransf.MoveFirst
    Do While Not ResTransf.EOF
        
        
            If ResTransf!nofi_tx_tipomov = "S" Then
                'Dados de NotasFiscais
                CboHistorico.Text = FunNulo(ResTransf!nofi_tx_historico)
                txtReferencia.Text = FunNulo(ResTransf!nofi_tx_ReferenciaTransferencia) '29/09/08 -Patrícia
                
                'Dados de Duplicatas
                DtpDtTransf.Value = Format(ResTransf!dupl_dt_vencimento, "dd/mm/yy")
                
                If ResTransf!dupl_tx_status = "C" Then
                    Status = 0
                ElseIf ResTransf!dupl_tx_status = "R" Then
                    Status = 1
                End If
                
                CboStatus.ListIndex = Status
                TxtVlTransf.Text = Format$(ResTransf!dupl_vl_valor, "standard")
                
                If Not IsNull(ResTransf!dupl_tx_tipopagrec) Then
                   For X = 0 To CboTipoPag.ListCount - 1
                       If CboTipoPag.ItemData(X) = CInt(ResTransf!dupl_tx_tipopagrec) Then
                           CboTipoPag.ListIndex = X
                       End If
                   Next
                End If
                
                TxtDocPagto.Text = FunNulo(ResTransf!dupl_tx_docpagrec)
                CboCCorrenteOrigem.BoundText = ResTransf!coco_cd_codigo
                
                
                MskContaOrigem.Text = ResTransf!plco_cd_CodReduzido
                CboPlanoContasOrigem.BoundText = ResTransf!plco_cd_CodReduzido
            Else
                CboCCorrenteDestino.BoundText = ResTransf!coco_cd_codigo
                'Dados de Distribuicao

                MskContaDestino.Text = ResTransf!plco_cd_CodReduzido
                CboPlanoContasDestino.BoundText = ResTransf!plco_cd_CodReduzido
            End If
            
     
        ResTransf.MoveNext
    Loop
    ResTransf.Close
End Sub

Function funChecarCampos() As Boolean
    funChecarCampos = True
    
    If FunObrigatorioCBOX(CboStatus, "O Status da transferência é obrigatório!") Then
        CboStatus.SetFocus
        Exit Function
    End If
    
    If FunObrigatorioTXT(TxtVlTransf, "O valor da transferência é obrigatório!") Then
        TxtVlTransf.SetFocus
        Exit Function
    End If
    
    If FunObrigatorioCBOX(CboTipoPag, "O tipo de pagamento da transferência é obrigatório!") Then
        CboTipoPag.SetFocus
        Exit Function
    End If
    
    If CboStatus.ItemData(CboStatus.ListIndex) = "2" And CboTipoPag.ItemData(CboTipoPag.ListIndex) = "1" Then
       If FunObrigatorioTXT(TxtDocPagto, "O documento de pagamento da transferência é obrigatório!") Then
            TxtDocPagto.SetFocus
            Exit Function
        End If
    End If
    
    If FunObrigatorioCBO(CboCCorrenteOrigem, "O banco origem é obrigatório!") Then
        CboCCorrenteOrigem.SetFocus
        Exit Function
    End If
    
    If FunObrigatorioCBO(CboCCorrenteDestino, "O banco destino é obrigatório!") Then
        CboCCorrenteDestino.SetFocus
        Exit Function
    End If
    
    
    
    funChecarCampos = False
End Function

Sub subGravaTransf()
    Dim XStatus As String
    Dim xNumFat As Long
    Dim XNumDoc As String
    Dim XMovimentacao As String
    Dim Cont As Integer
    Dim ResFormulario As Object
    Dim ResDoc As Object
    Dim Xdoc As String
    Dim xdatapag As String
    
    XStatus = funVerificaStatus(CboStatus.ListIndex)
    If XStatus = "C" Then
        xdatapag = ""
    Else
        xdatapag = Format(DtpDtTransf.Value, "dd/mm/yyyy")
    End If

    'Obter código de documento de pagto se não for definido um e se for realizado
    If TxtDocPagto.Text = "" And XStatus = "R" Then
        Xdoc = Format(CStr(FunPegaGuardaUltimo("PR")), "A00000")
    Else
        Xdoc = TxtDocPagto.Text
    End If
    
    'Obter o código do tipo de documento p/ transf.
    SubQOpenRecordset ResDoc, "SELECT tido_cd_tipodoc FROM TipoDocumentos WHERE tido_tx_sigla='TBCR'", "Estatico"
    XTipoDocComCusto = ResDoc!tido_cd_TipoDoc
    ResDoc.Close
    
    SubQOpenRecordset ResDoc, "SELECT tido_cd_tipodoc FROM TipoDocumentos WHERE tido_tx_sigla='TBCP'", "Estatico"
    XTipoDocSemCusto = ResDoc!tido_cd_TipoDoc
    ResDoc.Close
    
    XControle = FunPegaGuardaUltimo("C") 'Gera um nº de controle que será o mesmo para todos os registros desta gravação
    Cont = 0
    XMovimentacao = "S"
    
    While (Cont <= 1)
        'Insere o crédito e o débito na tabela de notas fiscais/duplicatas/distribuicao
    
        'Abre o banco de Dados
        SubQOpenRecordset ResFormulario, "SELECT * FROM NotasFiscais WHERE nofi_cd_notafiscal= -1", Dinamico
        ResFormulario.AddNew
                    
        'Gera novo Nº de documento e de fatura
        xNumFat = FunPegaGuardaUltimo("F")
        XNumDoc = Format(FunPegaGuardaUltimo("D"), "A000000000")
                    
        ResFormulario!nofi_tx_serie = "U"
        ResFormulario!nofi_tx_status = XStatus
        If XMovimentacao = "S" Then
            ResFormulario!tido_cd_TipoDoc = XTipoDocSemCusto
        Else
            ResFormulario!tido_cd_TipoDoc = XTipoDocComCusto
        End If
        ResFormulario!nofi_dt_entrada = Format$(DtpDtTransf.Value, "DD/MM/YY")
        ResFormulario!nofi_dt_emissao = Format$(DtpDtTransf.Value, "DD/MM/YY")
        ResFormulario!nofi_tx_historico = FunNulo(CboHistorico.Text)
        ResFormulario!nofi_vl_valor = FunNuloVal(TxtVlTransf.Text)
        ResFormulario!nofi_tx_tipomov = XMovimentacao
        ResFormulario!nofi_nr_fatura = xNumFat
        ResFormulario!nofi_nr_documento = XNumDoc
        ResFormulario!nofi_tx_distribuir = "N"
        ResFormulario!nofi_tx_tipofatura = "1"
        ResFormulario!nofi_tx_controle = XControle
        ResFormulario!empr_cd_empresa = PCodEmpresa
        ResFormulario!nofi_tx_NumCartaTransferencia = Null
        
        
        If (Len(Trim(txtReferencia.Text)) > 0) Then
            ResFormulario!nofi_tx_ReferenciaTransferencia = Trim(txtReferencia.Text)
        Else
            ResFormulario!nofi_tx_ReferenciaTransferencia = Null
        End If
        
        ResFormulario.Update
    
        Chave = ResFormulario!nofi_cd_notafiscal
            
        ResFormulario.Close
        
       If (XMovimentacao = "S") Then
            
            If Not (IsNull(CboPlanoContasOrigem.SelectedItem)) Then
                DatConta.Recordset.Bookmark = CboPlanoContasOrigem.SelectedItem
            End If
            
            tdbDistribuicao.MoveFirst
            While Not tdbDistribuicao.EOF
                Sql = "INSERT INTO Distribuicao (nofi_cd_notafiscal,plco_cd_conta,cere_cd_pcr,dist_vl_valor)" & _
                                  " VALUES (" & Chave & ",'" & DatConta.Recordset.Fields("plco_cd_conta") & "','" & tdbDistribuicao.Columns("CodPcr") & "'," & FunTrataFloat(tdbDistribuicao.Columns("Valor")) & ")"
                Conexao.Execute Sql
                tdbDistribuicao.MoveNext
            Wend
        Else
            
            If Not (IsNull(CboPlanoContasDestino.SelectedItem)) Then
                DatConta.Recordset.Bookmark = CboPlanoContasDestino.SelectedItem
            End If
            
            tdbDistribuicao.MoveFirst
            While Not tdbDistribuicao.EOF
                Sql = "INSERT INTO Distribuicao (nofi_cd_notafiscal,plco_cd_conta,cere_cd_pcr,dist_vl_valor)" & _
                                  " VALUES (" & Chave & ",'" & DatConta.Recordset.Fields("plco_cd_conta") & "','" & tdbDistribuicao.Columns("CodPcr") & "'," & FunTrataFloat(tdbDistribuicao.Columns("Valor")) & ")"
                Conexao.Execute Sql
                tdbDistribuicao.MoveNext
            Wend
        End If
        
        If Cont = 0 Then
           XMovimentacao = "D"
           XFL_FATURA = xNumFat
           
           Sql = "INSERT INTO Duplicatas (dupl_nr_duplicata,dupl_tx_debcred,dupl_tx_status,dupl_dt_vencimento,dupl_vl_valor," & _
              "dupl_dt_pagrec,dupl_tx_tipopagrec,dupl_tx_docpagrec,coco_cd_codigo,dupl_nr_fatura, dupl_tx_naorelfinanceiro)" & _
              " VALUES ('U','" & XMovimentacao & "','" & FunNulo(XStatus) & "'," & FunNuloData(DtpDtTransf.Value, NomeSgbd) & "," & FunNuloVal(FunTrataFloat(TxtVlTransf.Text)) & "," & FunNuloData(xdatapag, NomeSgbd) & "," & _
              "'" & CStr(CboTipoPag.ItemData(CboTipoPag.ListIndex)) & "','" & Xdoc & "'," & CboCCorrenteOrigem.BoundText & "," & xNumFat & ", '" & XLT_NAOSAIRELFINANCEIRO & "')"
           Conexao.Execute Sql
        
        Else
           XMovimentacao = "C"
           
           Sql = "INSERT INTO Duplicatas (dupl_nr_duplicata,dupl_tx_debcred,dupl_tx_status,dupl_dt_vencimento,dupl_vl_valor," & _
              "dupl_dt_pagrec,dupl_tx_tipopagrec,dupl_tx_docpagrec,coco_cd_codigo,dupl_nr_fatura, dupl_tx_naorelfinanceiro)" & _
              " VALUES ('U','" & XMovimentacao & "','" & FunNulo(XStatus) & "'," & FunNuloData(DtpDtTransf.Value, NomeSgbd) & "," & FunNuloVal(FunTrataFloat(TxtVlTransf.Text)) & "," & FunNuloData(xdatapag, NomeSgbd) & "," & _
              "'" & CStr(IIf(CboTipoPag.ItemData(CboTipoPag.ListIndex) = "4", "5", CboTipoPag.ItemData(CboTipoPag.ListIndex))) & "','" & Xdoc & "'," & CboCCorrenteDestino.BoundText & "," & xNumFat & ", '" & XLT_NAOSAIRELFINANCEIRO & "')"
            Conexao.Execute Sql
        End If
        
        Cont = Cont + 1
                
        '******************REGISTRA LOG DA OPERAÇÃO**************************
        ReDim XGM_MATRIZLOG(6, 2) As Variant
        
        XGM_MATRIZLOG(0, 0) = "Nº Doc" 'Número de documento da nota fiscal.
        XGM_MATRIZLOG(1, 0) = "PP" 'Número de fatura da nota fiscal.
        XGM_MATRIZLOG(2, 0) = "Status"  'Status da transferencia
        XGM_MATRIZLOG(3, 0) = "Movimentação"  'Informa se a nota é débito ou crédito
        XGM_MATRIZLOG(4, 0) = "Valor" 'Valor da nota fiscal.
        XGM_MATRIZLOG(5, 0) = "Historico" 'Histórico da nota fiscal
        XGM_MATRIZLOG(0, 1) = XNumDoc
        XGM_MATRIZLOG(1, 1) = xNumFat
        XGM_MATRIZLOG(2, 1) = XStatus
        XGM_MATRIZLOG(3, 1) = XMovimentacao
        XGM_MATRIZLOG(4, 1) = FunNuloVal(TxtVlTransf.Text)
        XGM_MATRIZLOG(5, 1) = FunNulo(CboHistorico.Text)
        
        Call subRegistraLog("frmTransferenciaComCusto", "1", funCriaDescricaoLog(XGM_MATRIZLOG, Transferencia, TRANSFERENCIA_BANCARIA))
        '*******************************************************************
        
        XMovimentacao = "E"
    Wend
    
    '***********CONTABILIDADE************
    If CboStatus.Text = "Realizado" Then
        SubQOpenRecordset ResFormulario, "SELECT dupl_cd_duplicata FROM Duplicatas WHERE dupl_nr_fatura = " & xNumFat & "", Estatico
        XFL_CODDUPL = ResFormulario!dupl_cd_Duplicata
        ResFormulario.Close
    End If
    '************************************
        
End Sub

Sub subPrepararInsercao()
    'Tab 0
    DtpDtTransf.Value = Format$(Now, "DD/MM/YY")
    CboStatus.ListIndex = 0
    TxtVlTransf.Text = ""
    CboTipoPag.ListIndex = 0
    CboCCorrenteOrigem.Text = ""
    CboCCorrenteDestino.Text = ""
    CboHistorico.Text = ""
    TxtDocPagto.Text = ""
    

End Sub

Function funVerificaStatus(Status As Integer) As String
    If Status = 0 Then 'Compromissado
            funVerificaStatus = "C"
    ElseIf Status = 1 Then 'Realizado
            funVerificaStatus = "R"
    End If
End Function

Private Sub CboCCorrenteOrigem_Change()
    TxtCCorrenteOrigem.Text = CboCCorrenteOrigem.BoundText

End Sub

Private Sub CboCCorrenteDestino_Change()
    TxtCCorrenteDestino.Text = CboCCorrenteDestino.BoundText
End Sub

Private Sub CboHistorico_GotFocus()
    CboHistorico.Text = "Transf. Bancária da C/C " & XContaOrigem & " para C/C " & XContaDestino
End Sub

Private Sub CboHistorico_KeyPress(KeyAscii As Integer)
    If Len(CboHistorico) = 100 Then
        If KeyAscii <> 8 And CboHistorico.SelLength = 0 Then
            KeyAscii = 0
            Beep
        End If
    End If
End Sub

Private Sub CboPlanoContasDestino_Click(Area As Integer)
    If CboPlanoContasDestino.Text <> "" Then
        DatConta.Recordset.Bookmark = CboPlanoContasOrigem.SelectedItem
        If DatConta.Recordset.Fields("plco_tx_tipo") <> "A" Then
            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a uma Transferência.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboPlanoContasDestino.BoundText = MskContaDestino.Text  '? Não deveria permitir gravar uma conta que não é analítica
            CboPlanoContasDestino.SetFocus
            Exit Sub
        End If
        MskContaDestino.Text = CboPlanoContasDestino.BoundText
        XCCDestino = DatConta.Recordset.Fields("plco_cd_conta")
    End If
End Sub

Private Sub CboPlanoContasOrigem_Change()
    If CboPlanoContasOrigem.Text <> "" Then
        DatConta.Recordset.Bookmark = CboPlanoContasOrigem.SelectedItem
        If DatConta.Recordset.Fields("plco_tx_tipo") <> "A" Then
            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a uma Tranferência.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboPlanoContasOrigem.BoundText = MskContaOrigem.Text  '? Não deveria permitir gravar uma conta que não é analítica
            CboPlanoContasOrigem.SetFocus
            Exit Sub
        End If
        MskContaOrigem.Text = CboPlanoContasOrigem.BoundText
        XCCOrigem = DatConta.Recordset.Fields("plco_cd_conta")
    End If
   
End Sub


Private Sub CboStatus_LostFocus()
    If (CboTipoPag.ListIndex > -1) Then
        If CboStatus.ListIndex = 1 And (CboTipoPag.ItemData(CboTipoPag.ListIndex)) <> "6" Then 'Status = Realizado e Tipo Pag <> Dinheiro
            TxtDocPagto.Enabled = True
        Else
            TxtDocPagto.Enabled = False
        End If
    End If
End Sub

Private Sub CboTipoPag_LostFocus()
    If (CboTipoPag.ListIndex > -1) Then
        If CboStatus.ListIndex = 1 And (CboTipoPag.ItemData(CboTipoPag.ListIndex)) <> "6" Then 'Status = Realizado e Tipo Pag <> Dinheiro
            TxtDocPagto.Enabled = True
        Else
            TxtDocPagto.Enabled = False
        End If
    End If
End Sub

Private Sub cmdAlterarContaCorrenteDestino_Click()
    If sstTab.Tab = 1 Then
        tdbDistribuicao.Enabled = False
        pnlDistribuicao.Visible = True
        pnlDistribuicao.Enabled = True
        XDistribuicao = "A"
        'Lucas Santiago - 27.01.09 - tratamento de variável nula ou vazia.
        If IsNull(tdbDistribuicao.Columns(1).Text) = False And tdbDistribuicao.Columns(1).Text <> "" Then
            txtCodigoCentroCustoDestino.Text = tdbDistribuicao.Columns(1).Text
        End If
        
        cboCentroCustoDestino.BoundText = txtCodigoCentroCustoDestino.Text
        XPcr = cboCentroCustoDestino.BoundText
        
        'Lucas Santiago - 27.01.09 - tratamento de variável nula ou vazia.
        If IsNull(tdbDistribuicao.Columns(2).Text) = False And tdbDistribuicao.Columns(2).Text <> "" Then
            txtValorDistribuicao.Text = Format(CDbl(tdbDistribuicao.Columns(2).Text), "standard")
        End If
        
        CmdGravar.Enabled = False
        
        If txtCodigoCentroCustoDestino.Enabled = True Then txtCodigoCentroCustoDestino.SetFocus
        
        Exit Sub
    End If

TrataErro:

  funTrataErros (ComMensagem)

End Sub

Private Sub cmdConfirmarContaCorrenteDestino_Click()
    Dim XLF_VALORNOTA As Double, XLF_VALORDISTRIB As Double
    Dim Xvalor
    Dim Mensagem As String
    Dim ResEstoque As Object
       
    If XDistribuicao = "A" Then
       If ValorTotalDistribuicao = 0 Then
       Else
            If CDbl(txtValorDistribuicao.Text) > 0 Then
                 ValorTotalDistribuicao = ValorTotalDistribuicao - CDbl(txtValorDistribuicao.Text)
            Else
                 ValorTotalDistribuicao = ValorTotalDistribuicao + CDbl(txtValorDistribuicao.Text)
            End If

       End If
    End If
    
    XLF_VALORNOTA = CDbl(LblValorNota.Caption)
    XLF_VALORDISTRIB = (ValorTotalDistribuicao + CDbl(txtValorDistribuicao.Text))
    If CDbl(Format(((CDbl(LblValorNota.Caption) * 100) / 100), "##,##0.00")) < CDbl(Format((((ValorTotalDistribuicao + CDbl(txtValorDistribuicao.Text)) * 100) / 100), "##,##0.00")) Then
        MsgBox "As distribuições excederam o valor total da Nota Fiscal", vbCritical, "ATENÇÃO"
        txtValorDistribuicao.SetFocus
        Exit Sub
    Else
        If (XPcr = txtCodigoCentroCustoDestino.Text) Then
            If XDistribuicao = "I" Then
                subInsereRegistroVetor Array(DatPCR.Recordset.Fields("cere_cd_pcr"), cboCentroCustoDestino.BoundText, txtValorDistribuicao.Text), arrDistribuicao, tdbDistribuicao
                tdbDistribuicao.MoveFirst
            ElseIf XDistribuicao = "A" Then
''''''                    If CDbl(txtValorDistribuicao.Text) > XLF_VALORNOTA Then
''''''                        MsgBox "As distribuições excederam o valor total da Nota Fiscal", vbCritical, "ATENÇÃO"
''''''                        txtValorDistribuicao.SetFocus
''''''                        Exit Sub
''''''                    Else
''''''                        tdbDistribuicao.Columns(0).Text = DatPCR.Recordset.Fields("cere_cd_pcr")
''''''                        tdbDistribuicao.Columns(1).Text = cboCentroCustoDestino.BoundText
''''''                        tdbDistribuicao.Columns(2).Text = txtValorDistribuicao.Text
''''''                    End If
                    
                    
                        tdbDistribuicao.Columns(0).Text = DatPCR.Recordset.Fields("cere_cd_pcr")
                        tdbDistribuicao.Columns(1).Text = cboCentroCustoDestino.BoundText
                        tdbDistribuicao.Columns(2).Text = txtValorDistribuicao.Text

                                        
            End If
        Else
            If XDistribuicao <> "A" And PrimeiraInsercao = True Then
                
                    If XDistribuicao = "I" Then
                       subInsereRegistroVetor Array(DatPCR.Recordset.Fields("cere_cd_pcr"), cboCentroCustoDestino.BoundText, txtValorDistribuicao.Text), arrDistribuicao, tdbDistribuicao
                    ElseIf XDistribuicao = "A" Then
                        tdbDistribuicao.Columns(0).Text = DatPCR.Recordset.Fields("cere_cd_pcr")
                        tdbDistribuicao.Columns(1).Text = txtCodigoCentroCustoDestino.Text
                        tdbDistribuicao.Columns(2).Text = txtValorDistribuicao.Text
                    End If
            Else
                    If ExisteDistribuicao(1, Array(0, 2), Array(DatPCR.Recordset.Fields("cere_cd_pcr"), DatConta.Recordset.Fields("plco_cd_conta")), arrDistribuicao, tdbDistribuicao) Then
                       MsgBox "Distribuição já Cadastrada", vbCritical, "ATENÇÃO"
                       If XDistribuicao = "I" Then
                          txtValorDistribuicao.SetFocus
                       ElseIf XDistribuicao = "A" Then
                            ValorTotalDistribuicao = ValorTotalDistribuicao + CDbl(txtValorDistribuicao)
                            sstTab.Enabled = True
                            Exit Sub
                       End If
                    Else
                        If XDistribuicao = "I" Then
                           subInsereRegistroVetor Array(DatPCR.Recordset.Fields("cere_cd_pcr"), cboCentroCustoDestino.BoundText, txtValorDistribuicao.Text), arrDistribuicao, tdbDistribuicao
                        ElseIf XDistribuicao = "A" Then
                            tdbDistribuicao.Columns(0).Text = DatPCR.Recordset.Fields("cere_cd_pcr")
                            tdbDistribuicao.Columns(1).Text = txtCodigoCentroCustoDestino.Text
                            tdbDistribuicao.Columns(2).Text = txtValorDistribuicao.Text
                            
                        End If
                    End If
    
            End If

        End If
    End If
    
    pnlDistribuicao.Enabled = False
    pnlDistribuicao.Visible = False
    ValorTotalDistribuicao = 0
    tdbDistribuicao.MoveFirst
    
    While Not tdbDistribuicao.EOF
        ValorTotalDistribuicao = ValorTotalDistribuicao + CDbl(tdbDistribuicao.Columns(2).Text)
        tdbDistribuicao.MoveNext
    Wend
    
    If ValorTotalDistribuicao > XLF_VALORNOTA Then
        MsgBox "As distribuições excederam o valor total da Nota Fiscal", vbCritical, "ATENÇÃO"
    End If


    tdbDistribuicao.MoveFirst
    tdbDistribuicao.Enabled = True
    
    If XInserir = "I" Or XDistribuido = "S" Then
        If LblValorNota.Caption = CDbl(LblValorTotalDistrib.Caption) Then
           XDistribuido = "S"
        Else
           XDistribuido = "N"
        End If
    End If
    
    LblValorTotalDistrib.Caption = Format(ValorTotalDistribuicao, "##,##0.00")

    subConectarControleDadosNV DatPCR, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY cere_cd_estruturado", Estatico

    If EntSaida = "E" Then
        subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') ORDER BY plco_tx_desccontabil", Estatico
    ElseIf EntSaida = "S" Then
        subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') ORDER BY plco_tx_desccontabil", Estatico
    End If
    '----------------------------------------------------------
    CmdGravar.Enabled = True
    PrimeiraInsercao = False
    
'End If
    
End Sub

Private Sub CmdDesistir_Click()
    Unload Me
End Sub

Private Sub cmdExcluirContaCorrenteDestino_Click()
    tdbDistribuicao.Enabled = False
   
    If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma exclusão
        subRemoveRegistroVetor arrDistribuicao, tdbDistribuicao
    End If

    LblValorTotalDistrib.Caption = funSomaColuna(arrDistribuicao, 2)
    If arrDistribuicao(0, 0) <> Empty Then
    Else
        cmdInserirContaCorrenteDestino.Enabled = True
        LblValorTotalDistrib.Caption = "0"
    End If
   
    tdbDistribuicao.Enabled = True
 
    Exit Sub
End Sub

Private Sub cmdInserirContaCorrenteDestino_Click()
XDistribuicao = "I"
    
    If sstTab.Tab = 1 Then
        frmTransferenciaComCusto.MousePointer = vbHourglass
        XDistribuicao = "I"
        PrepararInsercao
        If LblValorNota.Caption = LblValorTotalDistrib Then
           MsgBox "As distribuições já atingiram o valor total da Nota Fiscal", vbCritical, "ATENÇÃO"
        Else
            pnlDistribuicao.Visible = True
            pnlDistribuicao.Enabled = True
            txtCodigoCentroCustoDestino.SetFocus
        End If
        
        frmTransferenciaComCusto.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdRetornarContaCorrenteDestino_Click()
    pnlDistribuicao.Visible = False
    pnlDistribuicao.Enabled = False
    XFT_INSERIR = ""
    tdbDistribuicao.Enabled = True
    
    If arrDistribuicao(0, 0) <> Empty Then
    Else
       cmdInserirContaCorrenteDestino.Enabled = True
    End If
    
    ValorTotalDistribuicao = CDbl(LblValorTotalDistrib.Caption)
End Sub

Private Sub sstTab_Click(PreviousTab As Integer)
    If sstTab.Tab = 1 And TxtVlTransf.Text = "" Then
          MsgBox "O valor da Transferência deverá ser informado.", vbCritical, "ATENÇÃO"
    End If
   
    If sstTab.Tab = 1 And Not EntrouDist Then
         If XGB_AlteraTransfCusto = False Then
                Sql = "SELECT cere_cd_Pcr,cere_cd_estruturado,ValorDistrib FROM ConsFINDistribNota WHERE nofi_cd_notafiscal = " & Chave & ""
                subCarregaVetor Sql, Array("cere_cd_Pcr", "cere_cd_estruturado", "ValorDistrib"), Vetor3, tdbDistribuicao
                EntrouDist = True
                If Vetor3(0, 0) = Empty Then
                    cmdInserirContaCorrenteDestino.Enabled = True
                End If
                LblValorNota.Caption = Format$(TxtVlTransf.Text, "##,##0.00")
                LblValorTotalDistrib.Caption = funSomaColuna(Vetor3, 2)
                ValorTotalDistribuicao = CDbl(LblValorTotalDistrib.Caption)
         '---Else
         End If
     End If
End Sub

Private Sub txtCodigoCentroCustoDestino_GotFocus()
    txtCodigoCentroCustoDestino.SelStart = 0
    If txtCodigoCentroCustoDestino.MaxLength > 0 Then
        txtCodigoCentroCustoDestino.SelLength = txtCodigoCentroCustoDestino.MaxLength
    Else
        txtCodigoCentroCustoDestino.SelLength = Len(Trim(txtCodigoCentroCustoDestino.Text))
    End If
End Sub

Private Sub txtCodigoCentroCustoDestino_LostFocus()
    txtCodigoCentroCustoDestino.PromptInclude = False
    If txtCodigoCentroCustoDestino.Text <> "" Then
        If Len(txtCodigoCentroCustoDestino.Text) <> Len(pPCRmascara) And pPCRnivel = 1 Then
            txtCodigoCentroCustoDestino.Text = funFormataCCusto(txtCodigoCentroCustoDestino.Text)
        End If
    
        txtCodigoCentroCustoDestino.Mask = pPCRmascara
        txtCodigoCentroCustoDestino.PromptInclude = True
        cboCentroCustoDestino.BoundText = txtCodigoCentroCustoDestino.Text
        If cboCentroCustoDestino.Text = "" Then
            MsgBox "Não existe item do Centro de Custo com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            If txtCodigoCentroCustoDestino.Visible = True Then txtCodigoCentroCustoDestino.SetFocus
        End If
    End If
    txtCodigoCentroCustoDestino.PromptInclude = True

End Sub

Private Sub cboCentroCustoDestino_Change()
    If cboCentroCustoDestino.Text <> "" Then
        DatPCR.Recordset.Bookmark = cboCentroCustoDestino.SelectedItem
        If DatPCR.Recordset.Fields("cere_tx_tipo") <> "A" Then
            MsgBox "Este item do Centro de Custo não é analítico, logo não pode ser atribuído a uma Transferência.", vbCritical + vbOKOnly, "ATENÇÃO"
            cboCentroCustoDestino.BoundText = txtCodigoCentroCustoDestino.Text
            cboCentroCustoDestino.SetFocus
            Exit Sub
        End If
        txtCodigoCentroCustoDestino.Text = cboCentroCustoDestino.BoundText
    End If
End Sub


Private Sub cmdGravar_Click()
    Dim ResNF As Object
    Dim RESDUPL As Object
    Dim Sql As String

   If LblValorNota.Caption <> LblValorTotalDistrib.Caption Then
        MsgBox "O Total das distribuições não confere com o valor da Transferência!", vbCritical, "ATENÇÃO"
        Exit Sub
   End If

     If CboCCorrenteOrigem.BoundText = CboCCorrenteDestino.BoundText Then
        MsgBox "Não é possível transferir para a mesma conta!", vbCritical, "ATENÇÃO"
        Exit Sub
    End If
    
    If funChecarCampos Then Exit Sub
    
    Conexao.BeginTrans
    
    '***************CONTABILIDADE*******************
    If XGB_PQUALICONT Then 'Se possui o módulo de contabilidade
        If funMesFechado(DtpDtTransf.Value) <> "A" Then
            MsgBox "Não é possível realizar a Transferência nessa data, pois o mês contábil já foi fechado. " _
                + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
            Conexao.CommitTrans
            Exit Sub
        End If
    End If
    '***********************************************
    
    '///////APAGAR TRANSFERÊNCIA BANCÁRIA///////////
    If XFormulario = "TelaNotasFiscais" Or XFormulario = "TelaFatura" Or XFormulario = "TelaDuplicatas" Then
        
        SubQOpenRecordset ResNF, "SELECT * FROM NotasFiscais WHERE nofi_tx_controle = '" & XControleTransf & "' and empr_cd_empresa = " & PCodEmpresa & "", "Estatico"
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
            Conexao.Execute "DELETE FROM Distribuicao WHERE nofi_cd_notafiscal=" & ResNF!nofi_cd_notafiscal
                                    
            ResNF.MoveNext
        Loop
        
        Conexao.Execute "DELETE FROM NotasFiscais WHERE nofi_tx_controle='" & XControleTransf & "'  and empr_cd_empresa = " & PCodEmpresa & ""
        ResNF.Close
    End If
    '//////////////////////////////////////////////////////
    
    Call subGravaTransf
    
    '***************CONTABILIDADE*******************
'    If CboStatus.Text = "Realizado" Then 'contabiliza só se for realizado, pois não existe provisão para TB
'        SubContabilizarParcela XFL_CODDUPL, DtpDtTransf, False, False
'    End If
    '***********************************************

    
    Conexao.CommitTrans
    
'    'Imprime PP
'    If XGT_IMPRIMEPP = "S" Then
'        If MsgBox("Deseja Imprimir Processo de Pagamento?", vbQuestion + vbYesNo, "ATENÇÃO") = vbYes Then
'            subImprimePP XFL_FATURA, False
'        End If
'    End If
    

    
    Unload Me
    XGB_AlteraTransfCusto = False
Exit Sub
ResumoDoErro:
    Unload Me
    XGB_AlteraTransfCusto = False
    Exit Sub

TrataErro:
    funTrataErros (ComMensagem)
    Conexao.RollbackTrans
    XGB_AlteraTransfCusto = False
    Resume ResumoDoErro
End Sub


Private Sub CmdLimparHistorico_Click()
    CboHistorico.BoundText = 0
    CboHistorico.Text = ""
End Sub



Private Sub Form_Activate()
    Set Formulario = frmTransferenciaComCusto
    If (XFormulario = "TelaNotasFiscais" Or XFormulario = "TelaDuplicatas") Then
        EntrouDist = False
    ElseIf XFormulario = "frmTransferenciaComCusto" Then
        EntrouDist = True
    End If
    
End Sub

Private Sub Form_Load()

    DtpDtTransf.Value = Now
    XDistribuido = "S"
    PrimeiraInsercao = True

    Dim ResTransfBanc As Object
    Dim XLO_RSEMPRESA As ADODB.Recordset ' Declaração em 20/10/08 - Patrícia
    ValorSomaVetor = 0
       
       
    arrDistribuicao.ReDim 0, 0, 0, 2
    arrDistribuicao.Clear
    tdbDistribuicao.Array = arrDistribuicao
    

    subManutencaoJanelasAtivas "I", "frmTransferenciaComCusto"
    
    subConectarControleDadosNV DatHistorico, "SELECT * FROM Historicos ORDER BY hist_tx_descricao", Estatico
    subConectarControleDadosNV DatPCR, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY cere_cd_estruturado", Estatico
    'Alteração filtro CC por CGC Raiz - Samuel 27/09/18
    If Mid(PEmpresa, 1, 10) = "QUALIDADOS" Or Mid(PEmpresa, 1, 8) = "QUALINFO" Then
        subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_CGCRaiz = '" & Mid(XGT_CGC, 1, 10) & "' ORDER BY coco_tx_Descricao", Estatico
    Else
        subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo where empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
    End If
    subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_Tipo = 'A' ORDER BY plco_tx_desccontabil", Estatico
    subConectarControleDadosNV DatBanco, "SELECT * FROM Bancos", Estatico
    
    
    CmdGravar.Caption = "&Gravar"

    'Combo do tipo de pagto/rec
    'Tab 0
    CboTipoPag.RemoveItem (1)
    CboTipoPag.RemoveItem (1)
    CboTipoPag.RemoveItem (1)
    CboTipoPag.RemoveItem (2)

    
    If XFormulario = "TelaNotasFiscais" Or XFormulario = "TelaFatura" Or XFormulario = "TelaDuplicatas" Then
        subPreencheCampos
    ElseIf XFormulario = "frmTransferenciaComCusto" Then
        
        subPrepararInsercao 'Prepara os campo para a inserção
    End If
    
    If CboStatus.Text = "Realizado" Then

        SubQOpenRecordset ResTransfBanc, "SELECT * FROM ConsFINDuplicatas WHERE dupl_tx_DocPagRec = '" & TxtDocPagto.Text & _
                                            "' AND  dupl_dt_PagRec = " & FunNuloData(DtpDtTransf.Value, NomeSgbd) & _
                                            " AND coco_cd_codigo = " & TxtCCorrenteOrigem.Text, "Estatico"
        
        If ResTransfBanc.RecordCount > 1 Then
        
            MsgBox "Esta transferência pertence a um cheque consolidado e não poderá ser alterada! Cancele o cheque para efetuar a operação.", vbCritical + vbOKOnly, "ATENÇÃO"
            CmdGravar.Enabled = False
        'amiga, eu estou tão feliz! cosegui um novo emprego! estava louca para sair daqui :) Vou para o hospiatal S Rafael
            
        End If
        ResTransfBanc.Close
        
    End If
    
    If XGB_AlteraTransfCusto = True Then
         Sql = "SELECT cere_cd_Pcr,cere_cd_estruturado,ValorDistrib FROM ConsFINDistribNota WHERE nofi_cd_notafiscal = " & Chave & ""
         subCarregaVetor Sql, Array("cere_cd_Pcr", "cere_cd_estruturado", "ValorDistrib"), Vetor3, tdbDistribuicao
         EntrouDist = True
         If Vetor3(0, 0) = Empty Then
             'DesabilitaBotoes
             cmdInserirContaCorrenteDestino.Enabled = True
         End If
         LblValorNota.Caption = Format$(TxtVlTransf.Text, "##,##0.00")
         LblValorTotalDistrib.Caption = funSomaColuna(Vetor3, 2)
         ValorTotalDistribuicao = CDbl(LblValorTotalDistrib.Caption)
         PrimeiraAlteracao = True
         
         While Not tdbDistribuicao.EOF
            ValorSomaVetor = ValorSomaVetor + CDbl(tdbDistribuicao.Columns(2).Text)
            tdbDistribuicao.MoveNext
         Wend
         tdbDistribuicao.MoveFirst
         
    End If
    
    txtCodigoCentroCustoDestino.Mask = ""
    txtCodigoCentroCustoDestino.Text = ""
    txtCodigoCentroCustoDestino.Mask = pPCRmascara
    txtCodigoCentroCustoDestino.PromptInclude = True
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "frmTransferenciaComCusto"
End Sub

Private Sub MskContaOrigem_GotFocus()
    Call subSelecionaMSK(MskContaOrigem)
End Sub

Private Sub MskContaOrigem_LostFocus()
    If MskContaOrigem.Text <> "" Then
        CboPlanoContasOrigem.BoundText = MskContaOrigem.Text
        If CboPlanoContasOrigem.Text = "" Then
            MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
            MskContaOrigem.SetFocus
        End If
    End If
End Sub


Private Sub MskContaDestino_GotFocus()
    Call subSelecionaMSK(MskContaDestino)
End Sub

Private Sub MskContaDestino_LostFocus()
    If MskContaDestino.Text <> "" Then
        CboPlanoContasDestino.BoundText = MskContaDestino.Text
        If CboPlanoContasDestino.Text = "" Then
            MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
            MskContaDestino.SetFocus
        End If
    End If
End Sub



Private Sub TxtCCorrenteOrigem_GotFocus()
    Call subSelecionaTXT(TxtCCorrenteOrigem)
End Sub

Private Sub TxtCCorrenteOrigem_LostFocus()
    If TxtCCorrenteOrigem.Text <> "" Then
        CboCCorrenteOrigem.BoundText = TxtCCorrenteOrigem.Text
        If CboCCorrenteOrigem.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            TxtCCorrenteOrigem.SetFocus
        End If
    End If
End Sub

Private Sub TxtCCorrenteDestino_GotFocus()
    Call subSelecionaTXT(TxtCCorrenteDestino)
End Sub

Private Sub TxtCCorrenteDestino_LostFocus()
    If TxtCCorrenteDestino.Text <> "" Then
        CboCCorrenteDestino.BoundText = TxtCCorrenteDestino.Text
        If CboCCorrenteDestino.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            TxtCCorrenteDestino.SetFocus
        End If
    End If
End Sub

Private Sub TxtDocPagto_GotFocus()
    If CboTipoPag.ListIndex = 3 Or CboTipoPag.ListIndex = 4 Or CboTipoPag.ListIndex = 5 Then
        TxtDocPagto.Enabled = False
    Else
        TxtDocPagto.Enabled = True
    End If
End Sub

Private Sub TxtDocPagto_LostFocus()
    If CboTipoPag.ListIndex = 0 Then
        TxtDocPagto.Text = Format(TxtDocPagto, "0000000")
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
        txtValorDistribuicao.Text = Format$(((CDbl(TxtVlTransf.Text) * CDbl(TxtPercDistribuicao.Text)) / 100), "##,##0.00")
    End If
End Sub

Private Sub txtValorDistribuicao_GotFocus()
    Call subSelecionaTXT(txtValorDistribuicao)

End Sub

Private Sub txtValorDistribuicao_KeyPress(KeyAscii As Integer)
'''    If KeyAscii = 46 Then KeyAscii = 44
'''    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
'''        KeyAscii = 0
'''    End If

End Sub

Private Sub txtValorDistribuicao_LostFocus()
    If txtValorDistribuicao.Text <> "" Then
        TxtPercDistribuicao.Text = ((CDbl(txtValorDistribuicao.Text) / CDbl(TxtVlTransf.Text)) * 100)
    End If
End Sub


Private Sub TxtVlTransf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub TxtVlTransf_LostFocus()
    If IsNumeric(TxtVlTransf.Text) Then
        TxtVlTransf.Text = Format$(TxtVlTransf.Text, "##,##0.00")
    Else
        TxtVlTransf.Text = ""
    End If
    
    LblValorNota.Caption = Format$(TxtVlTransf.Text, "##,##0.00")
    LblValorTotalDistrib.Caption = funSomaColuna(arrDistribuicao, 2)

    ValorTotalDistribuicao = CDbl(LblValorTotalDistrib.Caption)
End Sub
Function ExisteDistribuicao(Numchaves As Byte, coluna As Variant, VALOR As Variant, Vetor As XArray, Grid As TDBGrid) As Boolean
    Dim Cont As Integer, Achou As Boolean
    Dim X As Integer, Y As Integer
    
    'Verifica se o registro está cadastrado no vetor
    
    ExisteDistribuicao = False
    
    While X <= Vetor.UpperBound(1)
        If Vetor(X, coluna(0)) = CStr(VALOR(0)) Then
            If Numchaves = 1 Then
                ExisteDistribuicao = True
                X = Vetor.UpperBound(1)
            Else
                Achou = True
                Y = 1
                While Achou And Y <= Numchaves - 1
                    'If coluna(Y) <> Valor(Y) Then
                    If Vetor(X, coluna(Y)) <> CStr(VALOR(Y)) Then
                        Achou = False
                    End If
                    Y = Y + 1
                Wend
                If Achou Then
                    ExisteDistribuicao = True
                    X = Vetor.UpperBound(1)
                End If
            End If
        End If
        X = X + 1
    Wend
End Function
Sub PrepararInsercao()
     
    If sstTab.Tab = 1 Then
        TxtPercDistribuicao.Text = ""
        If LblValorNota.Caption <> "" Then
            If LblValorTotalDistrib.Caption = "" Then
                LblValorTotalDistrib.Caption = "0"
            End If
            txtValorDistribuicao.Text = Format(CStr(CDbl(LblValorNota.Caption) - CDbl(LblValorTotalDistrib.Caption)), "standard")
            Call txtValorDistribuicao_LostFocus
        End If
        txtCodigoCentroCustoDestino.Mask = ""
        txtCodigoCentroCustoDestino.Text = ""
        txtCodigoCentroCustoDestino.Mask = pPCRmascara
        cboCentroCustoDestino.Text = ""
    End If
    
End Sub

