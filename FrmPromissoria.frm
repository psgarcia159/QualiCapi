VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmPromissoria 
   Caption         =   "Promissórias"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraPromissoria 
      Height          =   5835
      Left            =   0
      TabIndex        =   29
      Top             =   -60
      Width           =   9495
      Begin VB.OptionButton optJatoTintaOuLaser 
         Caption         =   "Jato Tinta/Laser"
         Height          =   195
         Left            =   4800
         TabIndex        =   59
         Top             =   5520
         Width           =   1575
      End
      Begin VB.OptionButton optMatricial 
         Caption         =   "Matricial"
         Height          =   195
         Left            =   3480
         TabIndex        =   58
         Top             =   5520
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox TxtObservacao2 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   57
         Top             =   4800
         Visible         =   0   'False
         Width           =   9135
      End
      Begin VB.TextBox TxtEstado 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8910
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1500
         Width           =   390
      End
      Begin VB.TextBox TxtJurosFin 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8310
         TabIndex        =   22
         Text            =   "0,0000"
         Top             =   4080
         Width           =   960
      End
      Begin VB.TextBox TxtValorIndexado 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4290
         TabIndex        =   25
         Text            =   "0,0000"
         Top             =   4440
         Width           =   1515
      End
      Begin VB.TextBox TxtValorReal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4290
         TabIndex        =   21
         Text            =   "0,00"
         Top             =   4020
         Width           =   1515
      End
      Begin VB.TextBox TxtPeriodicidade 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8430
         MaxLength       =   3
         TabIndex        =   19
         Text            =   "01"
         Top             =   3600
         Width           =   270
      End
      Begin VB.TextBox TxtParcelas 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6090
         MaxLength       =   3
         TabIndex        =   18
         Text            =   "001"
         Top             =   3600
         Width           =   390
      End
      Begin VB.TextBox TxtPlano 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1230
         MaxLength       =   2
         TabIndex        =   16
         Text            =   "01"
         Top             =   3600
         Width           =   270
      End
      Begin VB.TextBox TxtPrimeiraParcela 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3870
         MaxLength       =   3
         TabIndex        =   17
         Text            =   "001"
         Top             =   3600
         Width           =   390
      End
      Begin VB.TextBox TxtObservacao 
         Height          =   315
         Left            =   1230
         MaxLength       =   100
         TabIndex        =   15
         Top             =   3180
         Width           =   8055
      End
      Begin VB.TextBox TxtFavorecido 
         Height          =   315
         Left            =   1230
         MaxLength       =   50
         TabIndex        =   12
         Top             =   2325
         Width           =   8100
      End
      Begin VB.TextBox TxtPraca 
         Height          =   315
         Left            =   6930
         MaxLength       =   20
         TabIndex        =   14
         Top             =   2715
         Width           =   2385
      End
      Begin VB.TextBox TxtCidade 
         Height          =   315
         Left            =   4950
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1500
         Width           =   2445
      End
      Begin VB.TextBox TxtCliente 
         Height          =   315
         Left            =   2610
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   6720
      End
      Begin VB.TextBox TxtEndereco 
         Height          =   315
         Left            =   1230
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1095
         Width           =   8100
      End
      Begin VB.TextBox TxtBairro 
         Height          =   315
         Left            =   1230
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1500
         Width           =   2385
      End
      Begin VB.CheckBox chkClausula 
         Caption         =   "Cláusula de Reajuste"
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
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   26
         Top             =   4500
         Width           =   2295
      End
      Begin MSMask.MaskEdBox MskContrato 
         Height          =   315
         Left            =   1230
         TabIndex        =   0
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "AAAA.AAAA.##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MskCep 
         Height          =   315
         Left            =   1230
         TabIndex        =   9
         Top             =   1920
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##.###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MskTelResidencial 
         Height          =   315
         Left            =   4125
         TabIndex        =   10
         Top             =   1920
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   15
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MskTelComercial 
         Height          =   315
         Left            =   7590
         TabIndex        =   11
         Top             =   1920
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker DtpVenda 
         Height          =   315
         Left            =   4410
         TabIndex        =   3
         Top             =   660
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yy"
         Format          =   63307779
         CurrentDate     =   37917.6931481482
      End
      Begin MSComCtl2.DTPicker DtpRegistroContrato 
         Height          =   315
         Left            =   7830
         TabIndex        =   4
         Top             =   660
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yy"
         Format          =   63307779
         CurrentDate     =   37917.6933217593
      End
      Begin MSMask.MaskEdBox MskCnpj 
         Height          =   315
         Left            =   1230
         TabIndex        =   13
         Top             =   2760
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   18
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker DtpVencimento 
         Height          =   315
         Left            =   1230
         TabIndex        =   20
         Top             =   4020
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yy"
         Format          =   63307779
         CurrentDate     =   37180
      End
      Begin MSDataListLib.DataCombo CboIndexador 
         Bindings        =   "FrmPromissoria.frx":0000
         Height          =   315
         Left            =   1230
         TabIndex        =   23
         Top             =   4440
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "moed_tx_descricao"
         BoundColumn     =   "moed_cd_codmoeda"
         Text            =   "CboIndexador"
      End
      Begin MSAdodcLib.Adodc DatIndexador 
         Height          =   330
         Left            =   1230
         Top             =   4440
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
      Begin Threed.SSCommand CmdLimparIndexador 
         Height          =   315
         Left            =   3030
         TabIndex        =   24
         Top             =   4440
         Width           =   360
         _Version        =   65536
         _ExtentX        =   635
         _ExtentY        =   556
         _StockProps     =   78
         MouseIcon       =   "FrmPromissoria.frx":001B
         Picture         =   "FrmPromissoria.frx":046D
      End
      Begin MSMask.MaskEdBox MskCpfCnpj 
         Height          =   315
         Left            =   1230
         TabIndex        =   2
         Top             =   660
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   18
         PromptChar      =   "_"
      End
      Begin VB.Label LblCpfCnpj 
         Caption         =   "CPF / CNPJ:"
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
         TabIndex        =   56
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Juros Financeiros:"
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
         Left            =   6510
         TabIndex        =   55
         Top             =   4140
         Width           =   1665
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
         Height          =   195
         Left            =   3690
         TabIndex        =   54
         Top             =   4500
         Width           =   585
      End
      Begin VB.Label LblSimbolo 
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
         Left            =   5910
         TabIndex        =   53
         Top             =   4500
         Width           =   435
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
         Height          =   210
         Left            =   120
         TabIndex        =   52
         Top             =   4500
         Width           =   1035
      End
      Begin VB.Label LblValorReal 
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
         Left            =   3630
         TabIndex        =   51
         Top             =   4080
         Width           =   645
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
         Left            =   5910
         TabIndex        =   50
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label LblVencimnto 
         Alignment       =   1  'Right Justify
         Caption         =   "1º Venc.:"
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
         TabIndex        =   49
         Top             =   4080
         Width           =   1020
      End
      Begin VB.Label LblMeses 
         Alignment       =   1  'Right Justify
         Caption         =   "meses"
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
         Left            =   8790
         TabIndex        =   48
         Top             =   3660
         Width           =   510
      End
      Begin VB.Label Label1 
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
         Left            =   7110
         TabIndex        =   47
         Top             =   3660
         Width           =   1290
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
         Left            =   4650
         TabIndex        =   46
         Top             =   3660
         Width           =   1410
      End
      Begin VB.Label LblPrimeiraParcela 
         Alignment       =   1  'Right Justify
         Caption         =   "Nº da 1ª Parcela:"
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
         Left            =   2190
         TabIndex        =   45
         Top             =   3660
         Width           =   1620
      End
      Begin VB.Label LblObservacao 
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
         Height          =   195
         Left            =   60
         TabIndex        =   44
         Top             =   3240
         Width           =   1095
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
         Height          =   225
         Left            =   495
         TabIndex        =   43
         Top             =   3660
         Width           =   660
      End
      Begin VB.Label LblCnpj 
         Alignment       =   1  'Right Justify
         Caption         =   "CPF / CNPJ:"
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
         TabIndex        =   42
         Top             =   2820
         Width           =   1155
      End
      Begin VB.Label LblFavorecido 
         Alignment       =   1  'Right Justify
         Caption         =   "Favorecido:"
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
         TabIndex        =   41
         Top             =   2400
         Width           =   1035
      End
      Begin VB.Label LblRegistroContrato 
         Alignment       =   1  'Right Justify
         Caption         =   "Registro Contrato:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6090
         TabIndex        =   40
         Top             =   735
         Width           =   1635
      End
      Begin VB.Label LblVenda 
         Alignment       =   1  'Right Justify
         Caption         =   "Venda:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3630
         TabIndex        =   39
         Top             =   735
         Width           =   675
      End
      Begin VB.Label LblPraca 
         Alignment       =   1  'Right Justify
         Caption         =   "Praça:"
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
         Left            =   6120
         TabIndex        =   38
         Top             =   2745
         Width           =   705
      End
      Begin VB.Label LblFoneResidencial 
         Alignment       =   1  'Right Justify
         Caption         =   "Tel Resid.:"
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
         Left            =   3090
         TabIndex        =   37
         Top             =   1980
         Width           =   1005
      End
      Begin VB.Label LblFoneComercial 
         Alignment       =   1  'Right Justify
         Caption         =   "Tel Com.:"
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
         Left            =   6630
         TabIndex        =   36
         Top             =   1980
         Width           =   855
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
         Left            =   690
         TabIndex        =   35
         Top             =   1950
         Width           =   465
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
         Left            =   8265
         TabIndex        =   34
         Top             =   1560
         Width           =   660
      End
      Begin VB.Label LblCliente 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente:"
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
         Left            =   450
         TabIndex        =   33
         Top             =   300
         Width           =   705
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
         Left            =   570
         TabIndex        =   32
         Top             =   1560
         Width           =   585
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
         Left            =   270
         TabIndex        =   31
         Top             =   1110
         Width           =   885
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
         Left            =   4170
         TabIndex        =   30
         Top             =   1560
         Width           =   705
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
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
      TabIndex        =   28
      Top             =   5880
      Width           =   1035
   End
   Begin VB.CommandButton CmdImprimir 
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
      Left            =   2880
      TabIndex        =   27
      Top             =   5880
      Width           =   1035
   End
End
Attribute VB_Name = "FrmPromissoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim XFO_CLIENTE As New ADODB.Recordset

Private Sub ImprimeJatoTintaLaser()
    Dim XLO_IMPRESSORA As Printer
    Dim XLT_IMPRESSORA As String
    Dim XLT_SQL As String
    
    
    'Abre os recordsets e verificar se eles possuem dados
   
    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM consCAPFormularios WHERE titulo = '" & MskContrato.Text & "." & TxtPlano.Text & "." & TxtParcelas.Text & ".00' AND empr_cd_empresa = " & PCodEmpresa & "")
   
    If XGO_RSRELATORIO.EOF Then
        MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
        Me.MousePointer = vbDefault
        Exit Sub
    End If

    'Abre o relatório
    FrmVisRelatorios.Caption = "Promissórias"
    
    If UCase(PCliente) = "FRANISA" Then
        Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\NotaPromissoriaFranisa.rpt")
    ElseIf UCase(PCliente) = "COSTA ANDRADE" Then
        Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\NotaPromissoriaCA.rpt")
    ElseIf UCase(PCliente) = "ACTITUR" Then
        Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\NotaPromissoriaActitur.rpt")
    ElseIf UCase(PCliente) = "RJ" Then
        Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\NotaPromissoriaRJ.rpt")
    Else
        Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\NotaPromissoria.rpt")
    End If
    
    
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO, 3, 1


    If UCase(PCliente) <> "FRANISA" Then
    
        '********** Ativa os Relatórios
        With XGR_RELATORIO.FormulaFields

'             If Me.chkCondomino.Value = 1 Then
'                .GetItemByName("temcondomino").Text = Chr(34) + "sim" + Chr(34)
'             Else
'                .GetItemByName("temcondomino").Text = Chr(34) + "nao" + Chr(34)
'             End If

'             .GetItemByName("EmpresaCondomino").Text = Chr(34) + Me.txtCondomino.Text + Chr(34)
'             .GetItemByName("CpfCnpjCondomino").Text = Chr(34) + Me.txtCpfCnpj.Text + Chr(34)
             
'             If Me.chkImprimedataBase.Value = 1 Then
'                .GetItemByName("imprimedatabase").Text = Chr(34) + "sim" + Chr(34)
'             Else
'                .GetItemByName("imprimedatabase").Text = Chr(34) + "nao" + Chr(34)
'             End If
             
        End With
    End If
    
    FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
End Sub

Private Sub ImprimeMatricial()

      'Melhorias futuras:
      'Desconsiderar a margem superior e esquerda ao imprimir
      'Gravar no banco o tamanho da promissória
      Dim XLO_RSCONFIG As New ADODB.Recordset 'Recodset com as configurações dos campos que serão impressos
      Dim XLT_SQL As String 'SQL para abrir XLO_RSCONFIG
      Dim XLT_RECEBECAMPO As String
      Dim XLT_CAMPO As String
      Dim XLT_TITULO As String
      Dim XLD_VENCIMENTO As Date
      Dim XLT_FONT As String
      Dim XLB_Negrito As Boolean
      
      Dim Cont As Integer
      Dim XLI_CONT As Integer
      Dim XLO_IMPRESSORA As Printer
      Dim XLT_IMPRESSORA As String
      
      Dim XLI_ALTURA As Single 'Altura do papel
      Dim XLI_LARGURA As Single 'Largura da página
      
      Dim XLI_ESQUERDA As Single 'Margem mínima esquerda
      Dim XLI_TOPO As Single 'Margem mínima do topo
      
      Dim XLO_TAMANHOPAPEL As New ADODB.Recordset
      Dim LargPapelClausula As Long, AltPapelClausula As Long
      
      'Nome da impressora padrão
      XLT_IMPRESSORA = Printer.DeviceName
      
      
      XLT_FONT = "ARIAL"
      
      
      XLI_CONT = 0
      
      Printer.ScaleMode = vbCentimeters   'vbCentimeters
      Printer.Orientation = 1  'Portrait
         
      XLI_ESQUERDA = Printer.ScaleX(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX), vbPixels, vbCentimeters)
      XLI_TOPO = Printer.ScaleY(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY), vbPixels, vbCentimeters)
      
      SubQOpenRecordset XLO_TAMANHOPAPEL, "SELECT form_nr_Altura,form_nr_Largura FROM formularios WHERE form_tx_NomeCampo='Promissoria' AND empr_cd_Empresa = " & PCodEmpresa, Dinamico
      
      If Not IsNull(XLO_TAMANHOPAPEL!form_nr_altura) Then
        XLI_ALTURA = XLO_TAMANHOPAPEL!form_nr_altura
      Else
        XLI_ALTURA = 9.32
      End If
      
      If Not IsNull(XLO_TAMANHOPAPEL!form_nr_largura) Then
        XLI_LARGURA = XLO_TAMANHOPAPEL!form_nr_largura
      Else
        XLI_LARGURA = 20.45
      End If
      
      XLO_TAMANHOPAPEL.Close
      Set XLO_TAMANHOPAPEL = Nothing
      
      XLI_ALTURA = Printer.ScaleY(XLI_ALTURA, vbCentimeters, vbTwips) 'Seta a altura
      XLI_LARGURA = Printer.ScaleX(XLI_LARGURA, vbCentimeters, vbTwips) 'Seta a largura
         
      Printer.Height = XLI_ALTURA
      Printer.Width = XLI_LARGURA
      
      Set XLO_RSCONFIG = Nothing
      
      'RS com as configuração do que deve ser impresso - o tipo de doc = Promissoria e a empresa
      XLT_SQL = "SELECT * FROM formularios WHERE form_nr_imprime = -1 and form_tx_TipoDocumento = 'PR' AND empr_cd_Empresa = " & PCodEmpresa
      
      SubQOpenRecordset XLO_RSCONFIG, XLT_SQL, Estatico
      
      'XLS_VENCIMENTO = funExtenso2(Format(DtpVencimento.Value, "mm"))
      
      'XLS_VENCIMENTO = Day(DtpVencimento.Value) & " DE " & XLD_VENCIMENTO & " DE " & Year(DtpVencimento.Value)
      
      XLD_VENCIMENTO = DtpVencimento.Value
      
      For XLI_CONT = CInt(TxtPrimeiraParcela.Text) To CInt(TxtParcelas.Text)
        
        XLT_TITULO = TxtPlano.Text & "." & Format(XLI_CONT, "000") & "/" & Right(TxtParcelas.Text, 2)
        
        If XLI_CONT > 1 Then
          XLD_VENCIMENTO = DateAdd("m", 1 * CInt(TxtPeriodicidade.Text), XLD_VENCIMENTO)
        End If
        
        While Not (XLO_RSCONFIG.EOF)
          If Not (XLO_RSCONFIG!form_nr_negrito = -1) Then
            XLB_Negrito = False
          Else
            XLB_Negrito = True
          End If
          Printer.Font.Bold = XLB_Negrito
          Printer.Font.Size = XLO_RSCONFIG!form_nr_tamfonte
          Printer.Font.Name = XLT_FONT
          XLT_CAMPO = XLO_RSCONFIG!form_tx_nomecampo
                
          Select Case (XLT_CAMPO)
          'campos que nao estao na consulta ( datas e valores por extenso )
              Case "CodTitulo"
                XLT_RECEBECAMPO = XLT_TITULO
                
              Case "CodImovel"
                XLT_RECEBECAMPO = Left(MskContrato.Text, 9)
                
              Case "NomeCliente"
                XLT_RECEBECAMPO = TxtCliente.Text
                
              Case "CPF_CGC"
                XLT_RECEBECAMPO = MskCpfCnpj.Text
                
              Case "DataVenda"
                Select Case (XLO_RSCONFIG!form_tx_Descricao)
                  Case "Dia da venda - 2 Dígitos"
                    XLT_RECEBECAMPO = Format(DtpVenda.Value, " dd")
                  Case "Mês da venda - Extenso"
                    XLT_RECEBECAMPO = Format(DtpVenda.Value, " mmmm ")
                  Case "Ano da venda - 4 Dígitos"
                    XLT_RECEBECAMPO = Format(DtpVenda.Value, "yyyy ")
                  Case "Data da venda - Exenso"
                    XLT_RECEBECAMPO = Format(DtpVenda.Value, "dd \dia\s \do \mê\s \de mmmm \do a\no \de yyyy")
                  Case "Ano da venda - 2 Dígitos"
                    XLT_RECEBECAMPO = Format(DtpVenda.Value, "yy ")
                Case Else
                  XLT_RECEBECAMPO = FunNulo(DtpVenda.Value)
                End Select
                
              Case "DataRegCont"
                Select Case (XLO_RSCONFIG!form_tx_Descricao)
                  Case "Dia do Registro Cont. - 2 Dig"
                    XLT_RECEBECAMPO = Format(DtpRegistroContrato.Value, "dd")
                  Case "Mês do Registro Cont - Extenso"
                    XLT_RECEBECAMPO = Format(DtpRegistroContrato.Value, " mmmm ")
                  Case "Ano Registro Cont. - 4 Dígitos"
                    XLT_RECEBECAMPO = Format(DtpRegistroContrato.Value, " yyyy ")
                  Case "Data Registro Cont. - Extenso"
                    XLT_RECEBECAMPO = Format(DtpRegistroContrato.Value, "dd \dia\s \do \mê\s \de mmmm \do a\no \de yyyy")
                  Case "Ano Registro Cont. - 2 Dígitos"
                    XLT_RECEBECAMPO = Format(DtpRegistroContrato.Value, " yy ")
                Case Else
                  XLT_RECEBECAMPO = FunNulo(DtpRegistroContrato.Value)
                End Select
                
                
                Case "DataVenc"
                Select Case (XLO_RSCONFIG!form_tx_Descricao)
                  Case "Dia do Vencimento - 2 Dígitos"
                    XLT_RECEBECAMPO = Format(XLD_VENCIMENTO, " dd")
                  Case "Mês do Vencimento - Extenso"
                    XLT_RECEBECAMPO = UCase(Format(XLD_VENCIMENTO, "mmmm "))
                  Case "Ano do Vencimento - 4 Dígitos"
                    XLT_RECEBECAMPO = Format(XLD_VENCIMENTO, "yyyy ")
                  Case "Data do Vencimento - Exenso"
                    'RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "dd             mmmm             yyyy")
                    XLT_RECEBECAMPO = Format(XLD_VENCIMENTO, "dd \dia\s \do \mê\s \de mmmm \do a\no \de yyyy")
                  Case "Ano do Vencimento - 2 Dígitos"
                    XLT_RECEBECAMPO = Format(XLD_VENCIMENTO, "yy ")
                  Case "Data do Vencimento - Exenso 2"
                    XLT_RECEBECAMPO = UCase(Format(XLD_VENCIMENTO, "dd \de mmmm \de yyyy"))
                  Case "Dia do Vencimento - Extenso"
                    XLT_RECEBECAMPO = funExtenso2(Format(XLD_VENCIMENTO, "dd"))
                  Case Else
                    XLT_RECEBECAMPO = FunNulo(XLD_VENCIMENTO)
                End Select
              
                
              'Case "EndCorrespondencia"
               ' XLT_RECEBECAMPO = TxtEndereco.Text
                
              Case "EndCorrespondencia"
                XLT_RECEBECAMPO = TxtEndereco.Text & " " & TxtBairro.Text
                
              Case "Municipio_UF"
                XLT_RECEBECAMPO = TxtCidade.Text & " " & TxtEstado.Text
                
              Case "CEP"
                XLT_RECEBECAMPO = MskCep.Text
                
              Case "TelRes"
                  XLT_RECEBECAMPO = MskTelResidencial.Text
                  
              Case "TelCom"
                  XLT_RECEBECAMPO = MskTelComercial.Text
                  
              Case "NomeEmpresa"
                  XLT_RECEBECAMPO = TxtFavorecido.Text
                  
              Case "CGC"
                  XLT_RECEBECAMPO = MskCnpj.Text
                  
              Case "CidadeEstadoEmpresa"
                  XLT_RECEBECAMPO = TxtPraca.Text
                  
              Case "Observacao"
                  XLT_RECEBECAMPO = TxtObservacao.Text
              
    '          Case "DataVenc"
    '            Select Case (XLO_RSCONFIG!form_tx_Descricao)
    '              Case "Dia do Vencimento - 2 Dígitos"
    '                XLT_RECEBECAMPO = Format(XLD_VENCIMENTO, " dd")
    '              Case "Mês do Vencimento - Extenso"
    '                XLT_RECEBECAMPO = Format(XLD_VENCIMENTO, "mmmm ")
    '              Case "Ano do Vencimento - 4 Dígitos"
    '                XLT_RECEBECAMPO = Format(XLD_VENCIMENTO, "yyyy ")
    '              Case "Data do Vencimento - Exenso"
    '                XLT_RECEBECAMPO = Format(XLD_VENCIMENTO, "dd \dia\s \do \mê\s \de mmmm \do a\no \de yyyy")
    '              Case "Ano do Vencimento - 2 Dígitos"
    '                XLT_RECEBECAMPO = Format(XLD_VENCIMENTO, "yy ")
    '              Case Else
    '                XLT_RECEBECAMPO = FunNulo(XLD_VENCIMENTO)
    '            End Select
    '
              Case "ValorTitulo"
                If Not (XLO_RSCONFIG!form_tx_Descricao = "Valor do Título") Then
                  XLT_RECEBECAMPO = funExtenso(TxtValorReal.Text)
                    If chkClausula = 1 Then
                      Printer.CurrentY = Cont + XLO_RSCONFIG!form_nr_altura
                      Printer.CurrentX = XLO_RSCONFIG!form_nr_largura
                      XLT_RECEBECAMPO = XLT_RECEBECAMPO & " REAJUSTADOS CONFORME O PACTUADO EM CONTRATO"
                      XLT_RECEBECAMPO = XLT_RECEBECAMPO & " *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** "
                      subQuebraTexto XLT_RECEBECAMPO, XLO_RSCONFIG!form_nr_largura, 15, XLO_RSCONFIG!form_nr_tamfonte, XLT_FONT, XLO_RSCONFIG!form_nr_negrito, False
                    Else
                      Printer.CurrentY = XLO_RSCONFIG!form_nr_altura
                      Printer.CurrentX = XLO_RSCONFIG!form_nr_largura
                      XLT_RECEBECAMPO = XLT_RECEBECAMPO & " *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** ***  *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** ***"
                      subQuebraTexto XLT_RECEBECAMPO, XLO_RSCONFIG!form_nr_largura, 15, XLO_RSCONFIG!form_nr_tamfonte, XLT_FONT, XLO_RSCONFIG!form_nr_negrito, False
                      LargPapelClausula = XLO_RSCONFIG!form_nr_largura
                      AltPapelClausula = Cont + XLO_RSCONFIG!form_nr_altura + 1.5
                    End If
                Else
                  XLT_RECEBECAMPO = TxtValorReal.Text
                End If
              
              Case "ValorIndexado"
                XLT_RECEBECAMPO = TxtValorIndexado.Text
                
              Case "JurosFin"
                XLT_RECEBECAMPO = TxtJurosFin.Text
                
              Case "NomeIndexador"
                If LblSimbolo.Caption = "R$" Then
                  XLT_RECEBECAMPO = ""
                Else
                  XLT_RECEBECAMPO = LblSimbolo.Caption
                End If
                
              Case "Cotacao"
                If CDbl(TxtValorIndexado.Text) <> 0 And CboIndexador.BoundText <> "" Then
                  XLT_RECEBECAMPO = Format(CDbl(TxtValorReal.Text) / CDbl(TxtValorIndexado.Text), "##0.0000")
                Else
                  XLT_RECEBECAMPO = "0.000"
                End If
                
              Case Else
                XLT_RECEBECAMPO = ""
            
          End Select
          
          subImprimeTexto XLT_RECEBECAMPO, XLO_RSCONFIG!form_nr_largura, XLO_RSCONFIG!form_nr_altura, XLO_RSCONFIG!form_nr_tamfonte, XLT_FONT, XLO_RSCONFIG!form_nr_negrito, False
     
          XLO_RSCONFIG.MoveNext
            
        Wend
        
        If Mid(PEmpresa, 1, 7) = "FRANISA" Then 'COlocar esta condição nos padrões da empresa
            XLT_RECEBECAMPO = " EM MOEDA CORRENTE NACIONAL, REAJUSTADOS CONFORME O PACTUADO EM CONTRATO,"
            XLT_RECEBECAMPO = XLT_RECEBECAMPO & " NA PRAÇA DE SALVADOR - BA *** *** *** *** *** *** *** *** *** *** "
            'Printer.CurrentY = AltPapelClausula
            Printer.CurrentY = 5.5
            Printer.CurrentX = LargPapelClausula
            subQuebraTexto XLT_RECEBECAMPO, CSng(LargPapelClausula), 15, 8, XLT_FONT, 1, False
            subImprimeTexto XLT_RECEBECAMPO, LargPapelClausula, 5.5, 8, XLT_FONT, 1, False
        End If
        
        XLO_RSCONFIG.MoveFirst
        
      Printer.NewPage
      Next
      
      Printer.EndDoc
      XLO_RSCONFIG.Close
      Set XLO_RSCONFIG = Nothing
      
      'Retorna para a impressora padrão
      For Each XLO_IMPRESSORA In Printers
        If XLO_IMPRESSORA.DeviceName = XLT_IMPRESSORA Then
          Set Printer = XLO_IMPRESSORA
          Exit For
        End If
      Next
End Sub


Private Sub CmdImprimir_Click()
    
    Me.MousePointer = vbHourglass

    If optMatricial.Value = True Then
        ImprimeMatricial
    Else
        ImprimeJatoTintaLaser
    End If
    
    Me.MousePointer = vbDefault
End Sub

Private Sub MskContrato_LostFocus()
  
  Dim XLT_SQL As String
  
  TxtFavorecido.Text = PEmpresa
  MskCnpj.Text = XGT_CGC
  TxtPraca.Text = PCidade & "/" & XGT_ESTADO
  
  XLT_SQL = "SELECT * FROM ConsCAPPromissorias " & _
    " WHERE empr_cd_Empresa=" & PCodEmpresa & _
    " AND empd_cd_Empreendimento ='" & Left(MskContrato.Text, 4) & "'" & _
    " AND imov_cd_Imovel = '" & Mid(MskContrato.Text, 6, 4) & "'" & _
    " AND cont_cd_Contrato = '" & Right(MskContrato.Text, 2) & "'"
  
  SubQOpenRecordset XFO_CLIENTE, XLT_SQL, Estatico
    
  If Not XFO_CLIENTE.EOF Then
    
    'if XFO_CLIENTE!focl_tx_Tipo = "F" then
    
    TxtCliente.Text = FunNulo(XFO_CLIENTE!focl_tx_RazaoSocial)
    MskCpfCnpj.Text = FunNulo(XFO_CLIENTE!focl_tx_CGCCPF)
    TxtEndereco.Text = FunNulo(XFO_CLIENTE!clie_tx_endcorresp)
    TxtBairro.Text = FunNulo(XFO_CLIENTE!clie_tx_bairrocorresp)
    TxtCidade.Text = FunNulo(XFO_CLIENTE!clie_tx_muncorresp)
    TxtEstado.Text = FunNulo(XFO_CLIENTE!clie_tx_estcorresp)
    
    
    MskCep.PromptInclude = False
    MskCep.Text = FunNulo(XFO_CLIENTE!clie_nr_cepcorresp)
    MskCep.PromptInclude = True
    
    MskTelResidencial.Text = FunNulo(XFO_CLIENTE!focl_tx_fone1)
    MskTelComercial.Text = FunNulo(XFO_CLIENTE!focl_tx_fone)
        
    If Not IsNull(XFO_CLIENTE!cont_dt_RegistroContrato) Then
      DtpRegistroContrato.Value = XFO_CLIENTE!cont_dt_RegistroContrato
    End If
    
    If Not IsNull(XFO_CLIENTE!cont_dt_RegistroContrato) Then
      DtpVenda.Value = XFO_CLIENTE!cont_dt_Venda
    End If
    
    TxtJurosFin.Text = Format$(FunNuloVal(XFO_CLIENTE!cont_vl_PercJurosFin), "##0.0000")
  
  
  Else
    
    TxtCliente.Text = ""
    MskCpfCnpj.Text = ""
    TxtEndereco.Text = ""
    TxtBairro.Text = ""
    TxtCidade.Text = ""
    TxtEstado.Text = ""
    
    MskCep.PromptInclude = False
    MskCep.Text = ""
    MskCep.PromptInclude = True
    
    MskTelResidencial.Text = ""
    MskTelComercial.Text = ""
        
    DtpRegistroContrato.Value = Date
    DtpVenda.Value = Date
    DtpVencimento.Value = Date
    
    TxtJurosFin.Text = "0,0000"
    TxtValorReal.Text = "0,00"
    TxtValorIndexado.Text = "0,0000"
    
    CboIndexador.BoundText = ""
    'LblIndexador.Caption = ""
    chkClausula.Value = 0
    
    TxtPlano.Text = "01"
    TxtPrimeiraParcela.Text = "001"
    TxtParcelas.Text = "001"
    TxtPeriodicidade.Text = "01"
  
  End If
  
End Sub

Private Sub CboIndexador_Change()

    If CboIndexador.Text <> "" Then
      DatIndexador.Recordset.bookmark = CboIndexador.SelectedItem
      If Not IsNull(DatIndexador.Recordset.Fields("moed_tx_simbolo")) Then
        TxtValorIndexado.Text = FunCalcularValorIndexado(CDbl(TxtValorReal.Text), Format(DtpVenda.Value), CboIndexador.BoundText, DatIndexador.Recordset.Fields("moed_tx_Indexacao"), 0, TelaCotacoes)
        LblSimbolo.Caption = DatIndexador.Recordset.Fields("moed_tx_simbolo")
      End If
    End If

End Sub

Private Sub CmdCancelar_Click()
  
  Unload Me
  
End Sub

Private Sub CmdLimparIndexador_Click()
     
   CboIndexador.BoundText = ""
   LblSimbolo.Caption = "R$"
   TxtValorIndexado.Text = TxtValorReal.Text
  
End Sub

Private Sub Form_Load()
  
  subConectarControleDadosNV DatIndexador, "SELECT * FROM Moedas ORDER BY moed_tx_descricao", Estatico
  DtpVencimento.Value = Date
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  If XFO_CLIENTE.State = adStateOpen Then
    XFO_CLIENTE.Close
  End If
  Set XFO_CLIENTE = Nothing
  
End Sub




Private Sub TxtPeriodicidade_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub TxtPeriodicidade_LostFocus()
  
  If TxtPeriodicidade.Text = "" Then
    TxtPeriodicidade.Text = "000"
  Else
    TxtPeriodicidade.Text = Format(TxtPeriodicidade.Text, "00")
  End If
  
End Sub

Private Sub TxtPeriodicidade_GotFocus()
  
  Call subSelecionaTXT(TxtPeriodicidade)
  
End Sub

Private Sub TxtParcelas_KeyPress(KeyAscii As Integer)
  
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
  
End Sub

Private Sub TxtParcelas_LostFocus()
  
  If TxtParcelas.Text = "" Then
    TxtParcelas.Text = "000"
  Else
    TxtParcelas.Text = Format(TxtParcelas.Text, "000")
  End If
  
End Sub

Private Sub TxtParcelas_GotFocus()
  
  Call subSelecionaTXT(TxtParcelas)
  
End Sub

Private Sub TxtPrimeiraParcela_GotFocus()
  
  Call subSelecionaTXT(TxtPrimeiraParcela)
  
End Sub

Private Sub TxtPrimeiraParcela_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub TxtPrimeiraParcela_LostFocus()
  
  If TxtPrimeiraParcela.Text = "" Then
    TxtPrimeiraParcela.Text = "000"
  Else
    TxtPrimeiraParcela.Text = Format(TxtPrimeiraParcela.Text, "000")
  End If
  
End Sub

Private Sub TxtPlano_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub TxtPlano_LostFocus()
  
  If TxtPlano.Text = "" Then
    TxtPlano.Text = "00"
  Else
    TxtPlano.Text = Format(TxtPlano.Text, "00")
  End If
  
End Sub

Private Sub TxtPlano_GotFocus()
  
  Call subSelecionaTXT(TxtPlano)
 
End Sub

Private Sub TxtValorReal_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub TxtValorReal_LostFocus()
  
  If IsNumeric(TxtJurosFin.Text) Then
    TxtValorReal.Text = Format$(TxtValorReal.Text, "standard")
    If CboIndexador.BoundText = "" Then
      TxtValorIndexado.Text = TxtValorReal.Text
    End If
  Else
    TxtValorReal.Text = "0,00"
  End If
  
End Sub

Private Sub TxtValorReal_GotFocus()
  
  Call subSelecionaTXT(TxtValorReal)
  
End Sub

Private Sub TxtValorIndexado_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub TxtValorIndexado_LostFocus()
  
  If IsNumeric(TxtJurosFin.Text) Then
        'TxtValorIndexado.Text = Format$(TxtValorIndexado.Text, "##0.0000")
        TxtValorIndexado.Text = Format$(TxtValorIndexado.Text, "standard")
  Else
    TxtValorIndexado.Text = "0,0000"
  End If
  
End Sub

Private Sub TxtValorIndexado_GotFocus()
  
  Call subSelecionaTXT(TxtValorIndexado)
  
End Sub

Private Sub TxtJurosFin_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub TxtJurosFin_LostFocus()
    
    If IsNumeric(TxtJurosFin.Text) Then
        'TxtJurosFin.Text = Format$(TxtJurosFin.Text, "##0.0000")
        TxtJurosFin.Text = Format$(TxtJurosFin.Text, "standard")
    Else
        TxtJurosFin.Text = "0,0000"
    End If
    
End Sub
  
Private Sub TxtJurosFin_GotFocus()

  Call subSelecionaTXT(TxtJurosFin)

End Sub


    
