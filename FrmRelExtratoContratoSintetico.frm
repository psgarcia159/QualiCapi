VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRelExtratoContratoSintetico 
   Caption         =   "Relatório Extrato do Contrato - Sintético"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9345
   Icon            =   "FrmRelExtratoContratoSintetico.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4800
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   4275
      Left            =   4620
      TabIndex        =   29
      Top             =   -60
      Width           =   4695
      Begin VB.Frame Frame3 
         Height          =   675
         Left            =   60
         TabIndex        =   47
         Top             =   3540
         Width           =   4575
         Begin VB.CheckBox ChkLiquidadoAtualizado 
            Alignment       =   1  'Right Justify
            Caption         =   "Apresentar Liquidado Atualizado?"
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
            TabIndex        =   48
            Top             =   240
            Value           =   1  'Checked
            Width           =   4215
         End
      End
      Begin VB.Frame FraDataBase 
         Height          =   615
         Left            =   60
         TabIndex        =   44
         Top             =   120
         Width           =   4575
         Begin MSComCtl2.DTPicker DtpDtBase 
            Height          =   315
            Left            =   1140
            TabIndex        =   45
            Top             =   180
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MM/yy"
            Format          =   109510659
            CurrentDate     =   37180
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
            TabIndex        =   46
            Top             =   240
            Width           =   1050
         End
      End
      Begin VB.Frame FraTipoPlano 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   60
         TabIndex        =   33
         Top             =   780
         Width           =   4575
         Begin Threed.SSCommand CmdLimparTipoPlano 
            Height          =   315
            Left            =   4200
            TabIndex        =   13
            Top             =   240
            Width           =   315
            _Version        =   65536
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "FrmRelExtratoContratoSintetico.frx":2AFA
            Picture         =   "FrmRelExtratoContratoSintetico.frx":2B16
         End
         Begin MSDataListLib.DataCombo CboTiposPlanos 
            Bindings        =   "FrmRelExtratoContratoSintetico.frx":2C28
            Height          =   315
            Left            =   1740
            TabIndex        =   12
            Top             =   240
            Width           =   2445
            _ExtentX        =   4313
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
            Left            =   2100
            Top             =   240
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
         Begin VB.Label LblTipoPlano 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo do Plano:"
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
            Left            =   420
            TabIndex        =   34
            Top             =   300
            Width           =   1260
         End
      End
      Begin VB.Frame FraNaturezaPlano 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   60
         TabIndex        =   35
         Top             =   1500
         Width           =   4575
         Begin VB.ComboBox CboNaturezasPlanos 
            Height          =   315
            ItemData        =   "FrmRelExtratoContratoSintetico.frx":2C43
            Left            =   1740
            List            =   "FrmRelExtratoContratoSintetico.frx":2C59
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   240
            Width           =   2415
         End
         Begin Threed.SSCommand CmdLimparNaturezaPlano 
            Height          =   315
            Left            =   4185
            TabIndex        =   15
            Top             =   240
            Width           =   315
            _Version        =   65536
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "FrmRelExtratoContratoSintetico.frx":2C94
            Picture         =   "FrmRelExtratoContratoSintetico.frx":2CB0
         End
         Begin VB.Label LblNaturezaPlano 
            Caption         =   "Natureza do Plano:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   60
            TabIndex        =   36
            Top             =   300
            Width           =   1680
         End
      End
      Begin VB.Frame FraAtraso 
         Height          =   675
         Left            =   60
         TabIndex        =   37
         Top             =   2220
         Width           =   4575
         Begin VB.CheckBox ChkMulta 
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
            Height          =   315
            Left            =   2880
            TabIndex        =   16
            Top             =   180
            Width           =   795
         End
         Begin VB.CheckBox ChkJuros 
            Caption         =   "Juros"
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
            Left            =   3720
            TabIndex        =   17
            Top             =   180
            Width           =   795
         End
         Begin VB.Label LblMultaJuros 
            Caption         =   "Apresentar títulos vencidos com"
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
            TabIndex        =   38
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame FraPagos 
         Height          =   675
         Left            =   60
         TabIndex        =   39
         Top             =   2880
         Width           =   4575
         Begin VB.CheckBox ChkTitulosPagos 
            Alignment       =   1  'Right Justify
            Caption         =   "Apresentar títulos pagos com os acréscimos?"
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
            TabIndex        =   18
            Top             =   180
            Width           =   4215
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4275
      Left            =   0
      TabIndex        =   23
      Top             =   -60
      Width           =   4575
      Begin VB.Frame FraContrato 
         Height          =   615
         Left            =   60
         TabIndex        =   41
         Top             =   120
         Width           =   4455
         Begin MSMask.MaskEdBox MskContratoInicial 
            Height          =   315
            Left            =   1140
            TabIndex        =   0
            Top             =   180
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   12
            Mask            =   "AAAA.AAAA.##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MskContratoFinal 
            Height          =   315
            Left            =   2700
            TabIndex        =   1
            Top             =   180
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   12
            Mask            =   "AAAA.AAAA.##"
            PromptChar      =   " "
         End
         Begin VB.Label LblA 
            Caption         =   "a"
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
            TabIndex        =   43
            Top             =   240
            Width           =   195
         End
         Begin VB.Label LblContrato 
            Alignment       =   1  'Right Justify
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
            Height          =   225
            Left            =   240
            TabIndex        =   42
            Top             =   240
            Width           =   825
         End
      End
      Begin VB.Frame FraPagamento 
         Caption         =   "Período de Pagamento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   60
         TabIndex        =   30
         Top             =   1500
         Width           =   4470
         Begin MSComCtl2.DTPicker DtpInicioPagamento 
            Height          =   315
            Left            =   660
            TabIndex        =   4
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   109510659
            CurrentDate     =   37627.4196527778
         End
         Begin MSComCtl2.DTPicker DtpFimPagamento 
            Height          =   315
            Left            =   2760
            TabIndex        =   5
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   109510659
            CurrentDate     =   37627.4197685185
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Fim:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2220
            TabIndex        =   32
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Início:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame FraCorrecao 
         Caption         =   "Correção"
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
         Left            =   60
         TabIndex        =   27
         Top             =   2880
         Width           =   4455
         Begin VB.OptionButton OptValorHistorico 
            Caption         =   "Valor Histórico"
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
            TabIndex        =   9
            Top             =   300
            Width           =   1575
         End
         Begin VB.OptionButton OptIndexador1 
            Caption         =   "Indexador 1"
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
            Left            =   1680
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton OptIndexador2 
            Caption         =   "Indexador 2"
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
            Left            =   3000
            TabIndex        =   11
            Top             =   240
            Width           =   1395
         End
      End
      Begin VB.Frame FraPeriodo 
         Caption         =   "Período de Vencimento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   60
         TabIndex        =   24
         Top             =   780
         Width           =   4470
         Begin MSComCtl2.DTPicker DtpInicio 
            Height          =   315
            Left            =   660
            TabIndex        =   2
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   109510659
            CurrentDate     =   37459
         End
         Begin MSComCtl2.DTPicker DtpFim 
            Height          =   315
            Left            =   2760
            TabIndex        =   3
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   109510659
            CurrentDate     =   37627.4195601852
         End
         Begin VB.Label LblFim 
            Alignment       =   1  'Right Justify
            Caption         =   "Fim:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2220
            TabIndex        =   26
            Top             =   360
            Width           =   495
         End
         Begin VB.Label LblInicio 
            Alignment       =   1  'Right Justify
            Caption         =   "Início:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   495
         End
      End
      Begin Threed.SSFrame FraStatus 
         Height          =   660
         Left            =   60
         TabIndex        =   28
         Top             =   2220
         Width           =   4425
         _Version        =   65536
         _ExtentX        =   7805
         _ExtentY        =   1164
         _StockProps     =   14
         Caption         =   "Status do Contrato"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CheckBox ChkInativo 
            Caption         =   "Inativo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1740
            TabIndex        =   7
            Top             =   300
            Width           =   1080
         End
         Begin VB.CheckBox ChkAtivo 
            Caption         =   "Ativo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   360
            TabIndex        =   6
            Top             =   300
            Value           =   1  'Checked
            Width           =   1065
         End
         Begin VB.CheckBox ChkCedido 
            Caption         =   "Cedido"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3000
            TabIndex        =   8
            Top             =   300
            Width           =   1230
         End
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   8220
      TabIndex        =   22
      Top             =   4320
      Width           =   1035
   End
   Begin MSDataListLib.DataCombo CboObs 
      Bindings        =   "FrmRelExtratoContratoSintetico.frx":2DC2
      Height          =   315
      Left            =   1200
      TabIndex        =   19
      Top             =   4320
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
   Begin Threed.SSCommand CmdLimparObs 
      Height          =   315
      Left            =   6060
      TabIndex        =   20
      Top             =   4320
      Width           =   360
      _Version        =   65536
      _ExtentX        =   635
      _ExtentY        =   556
      _StockProps     =   78
      MouseIcon       =   "FrmRelExtratoContratoSintetico.frx":2DD7
      Picture         =   "FrmRelExtratoContratoSintetico.frx":3229
   End
   Begin MSAdodcLib.Adodc DatObs 
      Height          =   330
      Left            =   3600
      Top             =   4320
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
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir..."
      Height          =   345
      Left            =   6960
      TabIndex        =   21
      Top             =   4320
      Width           =   1035
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
      Left            =   0
      TabIndex        =   40
      Top             =   4380
      Width           =   1140
   End
End
Attribute VB_Name = "FrmRelExtratoContratoSintetico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim XFO_RSCLIENTESIMOVEIS As New ADODB.Recordset

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
    
    Dim XLT_FILTROPERIODO As String 'Armazenar filtro do período escolhido
    Dim XLT_FILTROCONTRATO As String
    Dim XLT_FILTROSTATUS As String
    Dim XLT_FILTROTIPOPLANO As String
    Dim XLT_FILTRONATUREZAPLANO As String
    Dim XLT_FILTROEMPREENDIMENTOS As String
    Dim XLT_FILTROTITULOSPAGOS As String
    Dim XLT_FILTROMULTA As String
    Dim XLT_FILTROJUROS As String
    Dim XLT_FILTROPERIODOPAGAMENTO As String
    Dim XLT_FILTROLIQUIDADOS As String
    Dim XLT_FILTROOBSERVACAO As String
    Dim XLO_TOTALINDEXADOR As Object
    Dim XLT_SQL As String
    
    Dim XLT_STATUS As String
    Dim XLT_STATUS2 As String
    Dim XLT_NATUREZA As String 'Natureza do Plano
    Dim XLT_INDEXADOR As String
    Dim XLT_ORDEM As String
    
    FrmRelExtratoContratoSintetico.MousePointer = vbHourglass
    
    XGT_SELECAO = ""
    XGT_CONJUNCAO = ""
    
    If NomeSgbd = "Access" Then
      
      XLT_SQL = "" & _
      "SELECT ConsCAPRelExtratoContrato.*," & _
      "IIF(Indexacao1='M'" & _
        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & "))" & _
      " as Cotacao1," & _
      "IIF(Indexacao2='M'" & _
        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & "))" & _
      "  as Cotacao2," & _
      "IIF(Indexacao3='M'" & _
        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=format(titu_dt_BasePagto,'\01/mm/yy'))" & _
        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=titu_dt_BasePagto))" & _
      " as Cotacao3,"
      
'      XLT_SQL = XLT_SQL & _
'      "IIF(Indexacao1='M'" & _
'        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")))" & _
'      " as Cotacao1_MesAnterior," & _
'      "IIF(Indexacao2='M'" & _
'        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")))" & _
'      " as Cotacao2_MesAnterior," & _
'      "IIF(Indexacao3='M'" & _
'        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=format(dateadd('m',-1,titu_dt_BasePagto),'\01/mm/yy')))" & _
'        "As Cotacao3_MesAnterior" & _
'      " From ConsCAPRelExtratoContrato WHERE "

      XLT_SQL = XLT_SQL & _
      "IIF(Indexacao1='M'" & _
        " ,IIF (cont_tx_CorrecProrata = 'A',(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
        " ,IIF (cont_tx_CorrecProrata = 'P',(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")))))" & _
      " as Cotacao1_MesAnterior," & _
      "IIF(Indexacao2='M'" & _
        " ,IIF (cont_tx_CorrecProrata = 'A',(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
        " ,IIF (cont_tx_CorrecProrata = 'P',(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")))))" & _
      " as Cotacao2_MesAnterior," & _
       "IIF(Indexacao3='M'" & _
        " ,IIF (cont_tx_CorrecProrata = 'A',(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=format(dateadd('m',-1,titu_dt_BasePagto),'\01/mm/yy'))" & _
        " ,IIF (cont_tx_CorrecProrata = 'P',(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=format(dateadd('m',1,titu_dt_BasePagto),'\01/mm/yy')))))" & _
      " as Cotacao3_MesAnterior" & _
      " From ConsCAPRelExtratoContrato WHERE "
    
    Else
    
      XLT_SQL = "" & _
      "SELECT ConsCAPRelExtratoContrato.*," & _
      "CASE Indexacao1   when 'M' then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
        " when 'D' then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & ")" & _
      " end as Cotacao1," & _
      "CASE  Indexacao2 when 'M' then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
        " when 'D'  then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & ")" & _
      " end as Cotacao2," & _
       "CASE  Indexacao3 when 'M' then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=convert(varchar,month(titu_dt_BasePagto))+'/01/' + convert(varchar,year(titu_dt_BasePagto)))" & _
        " when 'D'  then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=titu_dt_BasePagto)" & _
      " end as Cotacao3,"
      
'      XLT_SQL = XLT_SQL & _
'      "CASE  Indexacao1 when 'M' then" & _
'        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
'        " when 'D' then Null " & _
'      " end as Cotacao1_MesAnterior," & _
'      "CASE Indexacao2 when 'M' then" & _
'        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
'        " when 'D' then Null " & _
'      " end As Cotacao2_MesAnterior," & _
'      "CASE Indexacao3 when 'M' then" & _
'        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1,convert(varchar,month(titu_dt_BasePagto))+'/01/' + convert(varchar,year(titu_dt_BasePagto))))" & _
'        " when 'D' then Null " & _
'      " end As Cotacao3_MesAnterior" & _
'      " From ConsCAPRelExtratoContrato WHERE "

      XLT_SQL = XLT_SQL & _
      "CASE  Indexacao1 when 'M' then " & _
        "CASE cont_tx_CorrecProrata when 'A' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
        "when 'P' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
        "end " & _
      " when 'D' then Null " & _
      " end as Cotacao1_MesAnterior,"

      XLT_SQL = XLT_SQL & _
      "CASE Indexacao2 when 'M' then " & _
        "CASE cont_tx_CorrecProrata when 'A' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
        "when 'P' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
        "end " & _
      " when 'D' then Null " & _
      " end As Cotacao2_MesAnterior,"

      XLT_SQL = XLT_SQL & _
      "CASE Indexacao3 when 'M' then " & _
        "CASE cont_tx_CorrecProrata when 'A' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1,convert(varchar,month(titu_dt_BasePagto))+'/01/' + convert(varchar,year(titu_dt_BasePagto))))" & _
        "when 'P' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,1,convert(varchar,month(titu_dt_BasePagto))+'/01/' + convert(varchar,year(titu_dt_BasePagto))))" & _
        "end " & _
      " when 'D' then Null " & _
      " end As Cotacao3_MesAnterior" & _
      " From ConsCAPRelExtratoContrato WHERE "
    End If
  
     'Período de Vencimento
    If (Not IsNull(DtpInicio.Value)) And (Not IsNull(DtpFim.Value)) Then
       XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelExtratoContrato.titu_dt_Vencimento >= " & FunNuloData(DtpInicio.Value, NomeSgbd) & "" & _
        " AND ConsCAPRelExtratoContrato.titu_dt_Vencimento <= " & FunNuloData(DtpFim.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODO = "Vencimento: " & Format$(DtpInicio.Value, "dd/mm/yy") & " a " & Format$(DtpFim.Value, "dd/mm/yy")
    ElseIf (Not IsNull(DtpInicio.Value)) And (IsNull(DtpFim.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelExtratoContrato.titu_dt_Vencimento >= " & FunNuloData(DtpInicio.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODO = "Vencimento: A partir de " & Format$(DtpInicio.Value, "dd/mm/yy")
    ElseIf (IsNull(DtpInicio.Value)) And (Not IsNull(DtpFim.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "(ConsCAPRelExtratoContrato.titu_dt_Vencimento <= " & FunNuloData(DtpFim.Value, NomeSgbd) & ")"
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODO = "Vencimento: Até " & Format$(DtpFim.Value, "dd/mm/yy")
    Else
        XLT_FILTROPERIODO = "Vencimento: Geral"
    End If
    
    'Periodo de Pagamento
    If (Not IsNull(DtpInicioPagamento.Value)) And (Not IsNull(DtpFimPagamento.Value)) Then
       XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "(titu_dt_Pagamento is null OR (ConsCAPRelExtratoContrato.titu_dt_Pagamento >= " & FunNuloData(DtpInicioPagamento.Value, NomeSgbd) & "" & _
        " AND ConsCAPRelExtratoContrato.titu_dt_Pagamento <= " & FunNuloData(DtpFimPagamento.Value, NomeSgbd) & "))"
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODOPAGAMENTO = "Pagamento: " & Format$(DtpInicioPagamento.Value, "dd/mm/yy") & " a " & Format$(DtpFimPagamento.Value, "dd/mm/yy")
    ElseIf (Not IsNull(DtpInicioPagamento.Value)) And (IsNull(DtpFimPagamento.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "(titu_dt_Pagamento is null OR ConsCAPRelExtratoContrato.titu_dt_Pagamento >= " & FunNuloData(DtpInicioPagamento.Value, NomeSgbd) & ")"
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODOPAGAMENTO = "Pagamento: A partir de " & Format$(DtpInicioPagamento.Value, "dd/mm/yy")
    ElseIf (IsNull(DtpInicioPagamento.Value)) And (Not IsNull(DtpFimPagamento.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "(titu_dt_Pagamento is null OR ConsCAPRelExtratoContrato.titu_dt_Pagamento <= '" & Format(DtpFimPagamento.Value, "mm/dd/yyyy") & "')"
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODOPAGAMENTO = "Pagamento: Até " & Format$(DtpFimPagamento.Value, "dd/mm/yy")
    Else
        XLT_FILTROPERIODOPAGAMENTO = "Pagamento: Geral"
    End If
    
    'Contrato
    If MskContratoInicial.Text <> "    .    .  " And MskContratoFinal.Text <> "    .    .  " Then
        
        If Len(Trim(Replace(MskContratoInicial.Text, ".", ""))) = 4 Then
        
            XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "ConsCAPRelExtratoContrato.empd_cd_Empreendimento >= '" & Trim(Replace(MskContratoInicial.Text, ".", "")) & "' AND ConsCAPRelExtratoContrato.empd_cd_Empreendimento <= '" & Trim(Replace(MskContratoFinal.Text, ".", "")) & "'"
            XGT_CONJUNCAO = " AND "
            XLT_FILTROCONTRATO = "Contrato: " & MskContratoInicial & " a " & MskContratoFinal
            
        ElseIf Len(Trim(Replace(MskContratoInicial.Text, ".", ""))) = 12 Then
        
            XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "ConsCAPRelExtratoContrato.Contrato >= '" & Trim(Replace(MskContratoInicial.Text, ".", "")) & "' AND ConsCAPRelExtratoContrato.Contrato <= '" & Trim(Replace(MskContratoFinal.Text, ".", "")) & "'"
            XGT_CONJUNCAO = " AND "
            XLT_FILTROCONTRATO = "Contrato: " & MskContratoInicial & " a " & MskContratoFinal
        
        End If
        
      
    ElseIf MskContratoInicial.Text <> "    .    " And MskContratoFinal.Text = "    .    " Then
      
      If Len(Trim(Replace(MskContratoInicial.Text, ".", ""))) = 4 Then
      
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "ConsCAPRelExtratoContrato.empd_cd_Empreendimento >= '" & Trim(Replace(MskContratoInicial.Text, ".", "")) & "'"
        XGT_CONJUNCAO = " AND "
        XLT_FILTROCONTRATO = "Contrato: A partir de " & MskContratoInicial
    
      ElseIf Len(Trim(Replace(MskContratoInicial.Text, ".", ""))) = 12 Then
      
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "ConsCAPRelExtratoContrato.Contrato >= '" & MskContratoInicial.Text & "'"
        XGT_CONJUNCAO = " AND "
        XLT_FILTROCONTRATO = "Contrato: A partir de " & MskContratoInicial
      
      End If
    ElseIf MskContratoInicial.Text = "    .    " And MskContratoFinal.Text <> "    .    " Then
    
        If Len(Trim(Replace(MskContratoInicial.Text, ".", ""))) = 4 Then
      
            XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "ConsCAPRelExtratoContrato.empd_cd_Empreendimento <= '" & Trim(Replace(MskContratoFinal.Text, ".", "")) & "'"
            XGT_CONJUNCAO = " AND "
            XLT_FILTROCONTRATO = "Contrato: Até " & MskContratoInicial
            
        ElseIf Len(Trim(Replace(MskContratoInicial.Text, ".", ""))) = 12 Then
        
            XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "ConsCAPRelExtratoContrato.Contrato <= '" & MskContratoFinal.Text & "'"
            XGT_CONJUNCAO = " AND "
            XLT_FILTROCONTRATO = "Contrato: Até " & MskContratoInicial
            
        End If
      
    Else
      
      XLT_FILTROCONTRATO = "Contrato: Geral"
    
    End If
    
    'Status
    If ChkAtivo.Value = ChkInativo.Value And ChkAtivo.Value = ChkCedido.Value Then
        XLT_FILTROSTATUS = "Status: Geral"
        XLT_STATUS = "T" 'Todos
    Else
        If ChkAtivo.Value = 1 Then
            If ChkInativo.Value = 1 Then  'Ativo e Inativo
                XLT_STATUS = "A"
                XLT_STATUS2 = "I"
                XLT_FILTROSTATUS = "Status: Ativo e Inativo"
            ElseIf ChkCedido.Value = 1 Then  'Ativo e Reservado
                XLT_STATUS = "A"
                XLT_STATUS2 = "C"
                XLT_FILTROSTATUS = "Status: Ativo e Cedido"
            Else
                XLT_STATUS = "A" 'Ativo
                XLT_FILTROSTATUS = "Status: Ativo"
            End If
        ElseIf ChkInativo.Value = 1 Then
            If ChkCedido.Value = 1 Then   'Inativo e Reservado
                XLT_STATUS = "I"
                XLT_STATUS2 = "C"
                XLT_FILTROSTATUS = "Status: Inativo e Reservado"
            Else
                XLT_STATUS = "I" 'Inativo
                XLT_FILTROSTATUS = "Status: Inativo"
            End If
        ElseIf ChkCedido.Value = 1 Then
            XLT_STATUS = "C" 'Reservado
            XLT_FILTROSTATUS = "Status: Reservado"
        End If
        If XLT_STATUS2 <> "" Then 'Foi escolhido dois status
            XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "(ConsCAPRelExtratoContrato.cont_tx_Status = '" & XLT_STATUS & "'" & _
            " OR ConsCAPRelExtratoContrato.cont_tx_Status = '" & XLT_STATUS2 & "')"
            XGT_CONJUNCAO = " AND "
        Else
            XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "ConsCAPRelExtratoContrato.cont_tx_Status = '" & XLT_STATUS & "'"
            XGT_CONJUNCAO = " AND "
        End If
    End If
    
    'Tipo do Plano
    If CboTiposPlanos.BoundText <> "" Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO + "ConsCAPRelExtratoContrato.tipl_cd_TipoPlano = " & CboTiposPlanos.BoundText & ""
      XGT_CONJUNCAO = " AND "
      XLT_FILTROTIPOPLANO = "Tipo do Plano: " & CboTiposPlanos.Text
    Else
     XLT_FILTROTIPOPLANO = "Tipo do Plano: Geral"
    End If
        
    'Natureza
    If CboNaturezasPlanos.Text <> "" Then
      Select Case CboNaturezasPlanos.ListIndex
          Case 0
               XLT_NATUREZA = "S"
          Case 1
               XLT_NATUREZA = "P"
          Case 2
               XLT_NATUREZA = "F"
          Case 3
               XLT_NATUREZA = "G"
          Case 4
               XLT_NATUREZA = "O"
      End Select
     
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO + "ConsCAPRelExtratoContrato.titu_tx_NaturezaPlano = '" & XLT_NATUREZA & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTRONATUREZAPLANO = "Natureza do Plano: " & CboNaturezasPlanos.Text
    Else
      XLT_FILTRONATUREZAPLANO = "Natureza do Plano: Geral"
    End If
    
    'Indexador
    If OptValorHistorico.Value = True Then
      XLT_INDEXADOR = 0
    ElseIf OptIndexador1.Value = True Then
      XLT_INDEXADOR = 1
    ElseIf OptIndexador2.Value = True Then
      XLT_INDEXADOR = 2
    End If
    
    'Multa
    If ChkMulta.Value = 1 Then
      XLT_FILTROMULTA = "1"
    Else
      XLT_FILTROMULTA = "0"
    End If
    
    'Juros
    If ChkJuros.Value = 1 Then
      XLT_FILTROJUROS = "1"
    Else
      XLT_FILTROJUROS = "0"
    End If
    
    'Observacao
    If CboObs.BoundText <> "" Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsCAPRelExtratoContrato.obse_cd_Observacao=" & CboObs.BoundText & " "
      XGT_CONJUNCAO = " AND "
      XLT_FILTROOBSERVACAO = "Observação: " & CboObs.Text
    Else
      XLT_FILTROOBSERVACAO = "Observação:Geral"
    End If
    
    'Titulos pagos com atraso
    If ChkTitulosPagos.Value = 1 Then
      XLT_FILTROTITULOSPAGOS = "1"
    Else
      XLT_FILTROTITULOSPAGOS = "0"
    End If
    
    'Filtra a empresa Ativa
    XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + "ConsCAPRelExtratoContrato.empr_cd_Empresa = " & CStr(PCodEmpresa)
    
    'Fecha o recordset se estiver aberto
    If XGO_RSRELATORIO.State = adStateOpen Then
      XGO_RSRELATORIO.Close
      Set XGO_RSRELATORIO = Nothing
    End If
    'Abre os recordsets e verifica se eles possuem dados
    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute(XLT_SQL & XGT_SELECAO)
    
    If XGO_RSRELATORIO.EOF Then
        MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
        FrmRelExtratoContratoSintetico.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Fecha o recordset se estiver aberto
    If XGO_RSCABECALHO.State = adStateOpen Then
      XGO_RSCABECALHO.Close
      Set XGO_RSCABECALHO = Nothing
    End If
    Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
        
    'Fecha o recordset se estiver aberto
    If XFO_RSCLIENTESIMOVEIS.State = adStateOpen Then
      XFO_RSCLIENTESIMOVEIS.Close
      Set XFO_RSCLIENTESIMOVEIS = Nothing
    End If
    Set XFO_RSCLIENTESIMOVEIS = ConexaoRelatorio.Execute("SELECT * FROM ConsCapRelClientesImoveis WHERE empr_cd_empresa = " & PCodEmpresa)
     
    FrmVisRelatorios.Caption = "Relatório Extrato Contrato - Sintético"
        
    'Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelExtratoContratoSinteticoSemDem.rpt")
    Set XGR_RELATORIO = Nothing
    Set XGR_RELATORIO = RptExtratoContratoSintetico
    XGR_RELATORIO.DiscardSavedData
     
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
  
    'Ativa os Relatórios
    With XGR_RELATORIO.FormulaFields
        .GetItemByName("ApresentarLiquidadoAtualizado").Text = Chr(34) & (IIf(ChkLiquidadoAtualizado.Value = 1, "S", "N")) & Chr(34)
        .GetItemByName("FiltroContrato").Text = Chr(34) & XLT_FILTROCONTRATO & Chr(34)
        .GetItemByName("FiltroPeriodo").Text = Chr(34) & XLT_FILTROPERIODO & Chr(34)
        .GetItemByName("FiltroPeriodoPagamento").Text = Chr(34) & XLT_FILTROPERIODOPAGAMENTO & Chr(34)
        .GetItemByName("FiltroStatus").Text = Chr(34) & XLT_FILTROSTATUS & Chr(34)
        .GetItemByName("FiltroTipoPlano").Text = Chr(34) & XLT_FILTROTIPOPLANO & Chr(34)
        .GetItemByName("FiltroNaturezaPlano").Text = Chr(34) & XLT_FILTRONATUREZAPLANO & Chr(34)
        .GetItemByName("FiltroDataBase").Text = Chr(34) & Format(DtpDtBase.Value, "dd/mm/yy") & Chr(34)
        .GetItemByName("Indexador").Text = Chr(34) & XLT_INDEXADOR & Chr(34)
        .GetItemByName("FiltroMulta").Text = Chr(34) & XLT_FILTROMULTA & Chr(34)
        .GetItemByName("FiltroJuros").Text = Chr(34) & XLT_FILTROJUROS & Chr(34)
        .GetItemByName("FiltroObservacao").Text = Chr(34) & XLT_FILTROOBSERVACAO & Chr(34)
        .GetItemByName("FiltroTitulosPagos").Text = Chr(34) & XLT_FILTROTITULOSPAGOS & Chr(34)
        .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
    End With
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
    
    FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
    FrmRelExtratoContratoSintetico.MousePointer = vbDefault
End Sub

Private Sub CmdLimparNaturezaPlano_Click()
  
  CboNaturezasPlanos.ListIndex = -1
  
End Sub

Private Sub CmdLimparObs_Click()
    CboObs.Text = ""
End Sub

Private Sub CmdLimparTipoPlano_Click()
  
  CboTiposPlanos.BoundText = ""
  
End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelExtratoContratoSintetico
End Sub

Private Sub Form_Load()
  subManutencaoJanelasAtivas "I", "FrmRelExtratoContratoSintetico"
  subConectarControleDadosNV DatTipoPlano, "SELECT * FROM TiposPlanos ORDER BY tipl_tx_Descricao", Estatico
  subConectarControleDadosNV DatObs, "SELECT * FROM Observacoes ORDER BY obse_tx_observacao", Estatico
  DtpDtBase.Value = Date
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  If XGO_RSRELATORIO.State = adStateOpen Then
    XGO_RSRELATORIO.Close
  End If
  
  If XGO_RSCABECALHO.State = adStateOpen Then
    XGO_RSCABECALHO.Close
  End If
  
  If XGO_RSCABECALHO.State = adStateOpen Then
    XFO_RSCLIENTESIMOVEIS.Close
  End If
  
  Set XGO_RSRELATORIO = Nothing
  Set XGO_RSCABECALHO = Nothing
  Set XFO_RSCLIENTESIMOVEIS = Nothing
    
End Sub

Private Sub MskContratoInicial_GotFocus()
  
  Call subSelecionaMSK(MskContratoInicial)
  
End Sub

Private Sub MskContratoFinal_GotFocus()
  
  Call subSelecionaMSK(MskContratoFinal)
  
End Sub

Private Sub MskContratoInicial_KeyPress(KeyAscii As Integer)
  If KeyAscii >= 97 And KeyAscii <= 122 Then
    KeyAscii = KeyAscii - 32
  End If
End Sub

Private Sub MskContratoFinal_KeyPress(KeyAscii As Integer)
  If KeyAscii >= 97 And KeyAscii <= 122 Then
    KeyAscii = KeyAscii - 32
  End If
End Sub

Private Sub MskContratoInicial_LostFocus()
   
  If MskContratoInicial.Text <> "    .    .  " Then
    If Len(Trim(Replace(MskContratoInicial.Text, ".", ""))) <> 4 And Len(Trim(Replace(MskContratoInicial.Text, ".", ""))) <> 10 Then
      MsgBox "O código inicial do contrato deve possuir 4 ou 12 dígitos, ou deve estar vazio.", vbCritical, "ATENÇÃO !"
      MskContratoInicial.SetFocus
    End If
  End If
  
End Sub

Private Sub MskContratoFinal_LostFocus()
   
  If MskContratoFinal.Text <> "    .    .  " Then
    If Len(Trim(Replace(MskContratoFinal.Text, ".", ""))) <> 4 And Len(Trim(Replace(MskContratoFinal.Text, ".", ""))) <> 10 Then
      MsgBox "O código final do imóvel deve possuir 4 ou 12 dígitos, ou deve estar vazio", vbCritical, "ATENÇÃO !"
      MskContratoFinal.SetFocus
    End If
  End If
  
End Sub
