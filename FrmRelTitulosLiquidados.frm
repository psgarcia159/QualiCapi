VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRelTitulosLiquidados 
   Caption         =   "Títulos Liquidados"
   ClientHeight    =   4590
   ClientLeft      =   945
   ClientTop       =   2190
   ClientWidth     =   9390
   Icon            =   "FrmRelTitulosLiquidados.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkApresentarValorInicial 
      Alignment       =   1  'Right Justify
      Caption         =   "Apresentar Valor Original sem Correção?"
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
      Left            =   4800
      TabIndex        =   40
      Top             =   3120
      Width           =   4420
   End
   Begin VB.Frame FraContrato 
      Height          =   615
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   9375
      Begin MSMask.MaskEdBox MskContratoInicial 
         Height          =   315
         Left            =   3600
         TabIndex        =   36
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
         Left            =   5160
         TabIndex        =   37
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "AAAA.AAAA.##"
         PromptChar      =   " "
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
         Left            =   2700
         TabIndex        =   39
         Top             =   240
         Width           =   825
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
         Left            =   4980
         TabIndex        =   38
         Top             =   240
         Width           =   195
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3555
      Left            =   4680
      TabIndex        =   15
      Top             =   600
      Width           =   4695
      Begin VB.CheckBox ChkCorrecaoMonetaria 
         Alignment       =   1  'Right Justify
         Caption         =   "Correção Monetária Destacada?"
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
         Left            =   120
         TabIndex        =   29
         Top             =   3120
         Visible         =   0   'False
         Width           =   3255
      End
      Begin Threed.SSFrame FraStatus 
         Height          =   540
         Left            =   60
         TabIndex        =   24
         Top             =   120
         Width           =   4545
         _Version        =   65536
         _ExtentX        =   8017
         _ExtentY        =   952
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
            TabIndex        =   27
            Top             =   240
            Width           =   1230
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
            TabIndex        =   26
            Top             =   240
            Value           =   1  'Checked
            Width           =   1065
         End
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
            TabIndex        =   25
            Top             =   240
            Width           =   1080
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
         Height          =   540
         Left            =   60
         TabIndex        =   19
         Top             =   600
         Width           =   4575
         Begin Threed.SSCommand CmdLimparTipoPlano 
            Height          =   315
            Left            =   4200
            TabIndex        =   5
            Top             =   180
            Width           =   315
            _Version        =   65536
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "FrmRelTitulosLiquidados.frx":2AFA
            Picture         =   "FrmRelTitulosLiquidados.frx":2B16
         End
         Begin MSDataListLib.DataCombo CboTiposPlanos 
            Bindings        =   "FrmRelTitulosLiquidados.frx":2C28
            Height          =   315
            Left            =   1740
            TabIndex        =   4
            Top             =   180
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
            Top             =   180
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
            Left            =   60
            TabIndex        =   20
            Top             =   240
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
         Height          =   540
         Left            =   60
         TabIndex        =   22
         Top             =   1080
         Width           =   4575
         Begin VB.ComboBox CboNaturezasPlanos 
            Height          =   315
            ItemData        =   "FrmRelTitulosLiquidados.frx":2C43
            Left            =   1740
            List            =   "FrmRelTitulosLiquidados.frx":2C59
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   180
            Width           =   2415
         End
         Begin Threed.SSCommand CmdLimparNaturezaPlano 
            Height          =   315
            Left            =   4185
            TabIndex        =   7
            Top             =   180
            Width           =   315
            _Version        =   65536
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "FrmRelTitulosLiquidados.frx":2C94
            Picture         =   "FrmRelTitulosLiquidados.frx":2CB0
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
            Height          =   255
            Left            =   60
            TabIndex        =   23
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame frmTipoImovel 
         Height          =   540
         Left            =   60
         TabIndex        =   41
         Top             =   1560
         Width           =   4575
         Begin MSDataListLib.DataCombo CboTipoImovel 
            Bindings        =   "FrmRelTitulosLiquidados.frx":2DC2
            Height          =   315
            Left            =   1740
            TabIndex        =   43
            Top             =   120
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "tiim_tx_descricao"
            BoundColumn     =   "tiim_cd_TipoImovel"
            Text            =   "CboTipoImovel"
         End
         Begin Threed.SSCommand CmdLimpaTipoImovel 
            Height          =   315
            Left            =   4200
            TabIndex        =   44
            Top             =   120
            Width           =   315
            _Version        =   65536
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "FrmRelTitulosLiquidados.frx":2DDE
            Picture         =   "FrmRelTitulosLiquidados.frx":2DFA
         End
         Begin VB.Label lblTipoImovel 
            Caption         =   "Tipo Imóvel:"
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
            TabIndex        =   42
            Top             =   230
            Width           =   1095
         End
      End
      Begin VB.Frame FraAgrupados 
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
         Left            =   60
         TabIndex        =   21
         Top             =   2040
         Width           =   4575
         Begin VB.CheckBox chkJurosContratuais 
            Alignment       =   1  'Right Justify
            Caption         =   "Juros Cont"
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
            Height          =   375
            Left            =   3240
            TabIndex        =   28
            ToolTipText     =   "Mostrar juros contratuais"
            Top             =   120
            Width           =   1245
         End
         Begin VB.CheckBox ChkEmpreendimentos 
            Alignment       =   1  'Right Justify
            Caption         =   "Empreendimentos Agrupados?"
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
            Left            =   60
            TabIndex        =   8
            Top             =   120
            Value           =   1  'Checked
            Width           =   3015
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2950
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   4575
      Begin VB.Frame Frame3 
         Caption         =   "Período de Depósito"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   60
         TabIndex        =   30
         Top             =   1920
         Width           =   4470
         Begin MSComCtl2.DTPicker DtpInicioDeposito 
            Height          =   315
            Left            =   660
            TabIndex        =   31
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   82903043
            CurrentDate     =   37617.6519560185
         End
         Begin MSComCtl2.DTPicker DtpFimDeposito 
            Height          =   315
            Left            =   2760
            TabIndex        =   32
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   82903043
            CurrentDate     =   37617.6517939815
         End
         Begin VB.Label Label4 
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
            TabIndex        =   34
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label3 
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
            TabIndex        =   33
            Top             =   360
            Width           =   495
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
         Height          =   765
         Left            =   60
         TabIndex        =   16
         Top             =   1020
         Width           =   4470
         Begin MSComCtl2.DTPicker DtpInicioPagamento 
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
            Format          =   82903043
            CurrentDate     =   37617.6519560185
         End
         Begin MSComCtl2.DTPicker DtpFimPagamento 
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
            Format          =   82903043
            CurrentDate     =   37617.6517939815
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
            TabIndex        =   18
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
            TabIndex        =   17
            Top             =   360
            Width           =   495
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
         Height          =   765
         Left            =   60
         TabIndex        =   12
         Top             =   120
         Width           =   4470
         Begin MSComCtl2.DTPicker DtpInicio 
            Height          =   315
            Left            =   660
            TabIndex        =   0
            Top             =   285
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   82903043
            CurrentDate     =   37678
         End
         Begin MSComCtl2.DTPicker DtpFim 
            Height          =   315
            Left            =   2760
            TabIndex        =   1
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   82903043
            CurrentDate     =   37617.651087963
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
            TabIndex        =   14
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
            TabIndex        =   13
            Top             =   360
            Width           =   495
         End
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir..."
      Height          =   345
      Left            =   7005
      TabIndex        =   9
      Top             =   4200
      Width           =   1035
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   8160
      TabIndex        =   10
      Top             =   4200
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc DatTipoImovel 
      Height          =   330
      Left            =   4560
      Top             =   2280
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
      Caption         =   "DatTipoImovel"
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
Attribute VB_Name = "FrmRelTitulosLiquidados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChkApresentarValorInicial_Click()
    
    If ChkApresentarValorInicial.Value = 1 Then
        ChkCorrecaoMonetaria.Value = 0
        ChkCorrecaoMonetaria.Enabled = False
    Else
        ChkCorrecaoMonetaria.Enabled = True
    End If
    
End Sub

Private Sub ChkCorrecaoMonetaria_Click()
    
    If ChkCorrecaoMonetaria.Value = 1 Then
        ChkApresentarValorInicial.Value = 0
        ChkApresentarValorInicial.Enabled = False
    Else
        ChkApresentarValorInicial.Enabled = True
    End If
    
End Sub

Private Sub ChkEmpreendimentos_Click()
    If ChkEmpreendimentos.Value = 1 Then
        chkJurosContratuais.Enabled = False
        ChkApresentarValorInicial.Enabled = True
    Else
        chkJurosContratuais.Enabled = True
        ChkApresentarValorInicial.Enabled = False
        ChkApresentarValorInicial.Value = 0
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
    
    
    Dim XLT_FILTROSTATUS As String
    Dim XLT_FILTROTIPOPLANO As String
    Dim XLT_FILTRONATUREZAPLANO As String
    Dim XLT_FILTROTIPOIMOVEL As String
    Dim XLT_FILTROCONTRATO As String
    Dim XLT_FILTROTITULOSPAGOS As String
    Dim XLT_FILTROMULTA As String
    Dim XLT_FILTROJUROS As String
    Dim XLT_FILTRODESAGIO As String
    
    
    Dim XLT_FILTROPERIODOVENCIMENTO As String
    Dim XLT_FILTROPERIODODEPOSITO As String
    Dim XLT_FILTROPERIODOPAGAMENTO As String
    
    Dim XLT_FILTROLIQUIDADOS As String
    Dim XLT_FILTROOBSERVACAO As String
    
    Dim XLT_STATUS As String
    Dim XLT_STATUS2 As String
    Dim XLT_NATUREZA As String 'Natureza do Plano
    Dim XLT_TIPOIMOVEL As String 'Tipo Imóvel
    Dim XLT_INDEXADOR As String
    Dim XLT_ORDEM As String
    Dim XLT_SQL As String
    
    FrmRelTitulosLiquidados.MousePointer = vbHourglass
    
    XGT_SELECAO = ""
    XGT_CONJUNCAO = ""
    
    If NomeSgbd = "Access" Then
      
      XLT_SQL = "" & _
      "SELECT ConsCAPRelExtratoContrato.*," & _
      "IIF(Indexacao3='M'" & _
        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=format(titu_dt_BasePagto,'\01/mm/yy'))" & _
        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=titu_dt_BasePagto))" & _
      " as Cotacao3,"
      
'      XLT_SQL = XLT_SQL & _
'      "IIF(Indexacao3='M'" & _
'        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=format(dateadd('m',-1,titu_dt_BasePagto),'\01/mm/yy')))" & _
'        "As Cotacao3_MesAnterior" & _
'      " From ConsCAPRelExtratoContrato WHERE "
      
      XLT_SQL = XLT_SQL & _
      "IIF(Indexacao3='M'" & _
        " ,IIF (cont_tx_CorrecProrata = 'A',(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=format(dateadd('m',-1,titu_dt_BasePagto),'\01/mm/yy'))" & _
        " ,IIF (cont_tx_CorrecProrata = 'P',(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=format(dateadd('m',1,titu_dt_BasePagto),'\01/mm/yy')))))" & _
      " as Cotacao3_MesAnterior" & _
      " From ConsCAPRelExtratoContrato WHERE "
    
    Else
    
      XLT_SQL = "" & _
      "SELECT ConsCAPRelExtratoContrato.*," & _
       "CASE  Indexacao3 when 'M' then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=convert(varchar,month(titu_dt_BasePagto))+'/01/' + convert(varchar,year(titu_dt_BasePagto)))" & _
        " when 'D'  then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=titu_dt_BasePagto)" & _
      " end as Cotacao3,"
  
'      XLT_SQL = XLT_SQL & _
'      "CASE Indexacao3 when 'M' then" & _
'        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1,convert(varchar,month(titu_dt_BasePagto))+'/01/' + convert(varchar,year(titu_dt_BasePagto))))" & _
'        " when 'D' then Null " & _
'      " end As Cotacao3_MesAnterior" & _
'      " From ConsCAPRelExtratoContrato WHERE "

      XLT_SQL = XLT_SQL & _
      "CASE Indexacao3 when 'M' then " & _
        "CASE cont_tx_CorrecProrata when 'A' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1,convert(varchar,month(titu_dt_BasePagto))+'/01/' + convert(varchar,year(titu_dt_BasePagto))))" & _
        "when 'P' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,1,convert(varchar,month(titu_dt_BasePagto))+'/01/' + convert(varchar,year(titu_dt_BasePagto))))" & _
        "end " & _
      " when 'D' then Null " & _
      " end As Cotacao3_MesAnterior, "
      
      'Cotação na data da venda do imóvel
'      XLT_SQL = XLT_SQL & _
'      "CASE  Indexacao3 when 'M' then" & _
'        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=convert(varchar,month(cont_dt_Venda))+'/01/' + convert(varchar,year(cont_dt_Venda)))" & _
'        " when 'D'  then" & _
'        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=cont_dt_Venda)" & _
'      " end as Cotacao4,"
       
      'Cotação no inicio no plano
'      XLT_SQL = XLT_SQL & _
'      "(select " & _
'            "case when isnull(case titu_tx_IndicePagamento when 1 then titu_vl_ParcelaIndex1 else titu_vl_ParcelaIndex2 end, 0) <> 0 then titu_vl_Parcela/(case titu_tx_IndicePagamento when 1 then titu_vl_ParcelaIndex1 else titu_vl_ParcelaIndex2 end) else 0 end Cotacao " & _
'      " From ConsCAPRelExtratoContrato plano " & _
'      " Where plano.empr_cd_Empresa = ConsCAPRelExtratoContrato.empr_cd_Empresa " & _
'      " and right(plano.Titulo,6) = '001.00' " & _
'      " and left(plano.Titulo,15) =  left(ConsCAPRelExtratoContrato.Titulo,15))  Cotacao4, "
      XLT_SQL = XLT_SQL & _
        "(select top 1 " & _
            "case when ISNULL(CASE titu_tx_IndicePagamento WHEN 1 THEN titu_vl_ParcelaIndex1 ELSE titu_vl_ParcelaIndex2 END, 0) <> 0 then titu_vl_Parcela/(CASE titu_tx_IndicePagamento WHEN 1 THEN titu_vl_ParcelaIndex1 ELSE titu_vl_ParcelaIndex2 END) else 0 end Cotacao " & _
        "From ConsCAPRelExtratoContrato plano " & _
        "Where plano.empr_cd_Empresa = ConsCAPRelExtratoContrato.empr_cd_Empresa and left(plano.Titulo,15) =  left(ConsCAPRelExtratoContrato.Titulo,15) " & _
        "and plano.moed_cd_Moeda1 = ConsCAPRelExtratoContrato.moed_cd_Moeda1 order by Titulo asc)  Cotacao4,"

      XLT_SQL = XLT_SQL & _
      "CASE Indexacao3 when 'M' then " & _
        "CASE cont_tx_CorrecProrata when 'A' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1,convert(varchar,month(cont_dt_Venda))+'/01/' + convert(varchar,year(cont_dt_Venda))))" & _
        "when 'P' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,1,convert(varchar,month(cont_dt_Venda))+'/01/' + convert(varchar,year(cont_dt_Venda))))" & _
        "end " & _
      " when 'D' then Null " & _
      " end As Cotacao4_MesAnterior,  " & _
      " isnull(Residuo.ComResiduo,0) ComResiduo " & _
      " From ConsCAPRelExtratoContrato " & _
      " left join (" & _
      "            select " & _
      "                 ConsCAPRelExtratoContrato.empr_cd_Empresa CodEmp, " & _
      "                Substring(ConsCAPRelExtratoContrato.Titulo, 1, 19) Tit, " & _
      "                1 As ComResiduo " & _
      "            from ConsCAPRelExtratoContrato " & _
      "            group By ConsCAPRelExtratoContrato.empr_cd_Empresa, SUBSTRING(ConsCAPRelExtratoContrato.Titulo, 1,19)  " & _
      "            having count(Titulo) > 1 " & _
      " ) Residuo on Residuo.Tit =  substring(ConsCAPRelExtratoContrato.Titulo, 1,19) " & _
      "  and Residuo.CodEmp = ConsCAPRelExtratoContrato.empr_cd_Empresa " & _
      " where "
     
    End If
  
   'Período de Vencimento
    If (Not IsNull(DtpInicio.Value)) And (Not IsNull(DtpFim.Value)) Then
       XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelExtratoContrato.titu_dt_Vencimento >= " & FunNuloData(DtpInicio.Value, NomeSgbd) & "" & _
        " AND ConsCAPRelExtratoContrato.titu_dt_Vencimento <= " & FunNuloData(DtpFim.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODOVENCIMENTO = "Vencimento: " & Format$(DtpInicio.Value, "dd/mm/yy") & " a " & Format$(DtpFim.Value, "dd/mm/yy")
    ElseIf (Not IsNull(DtpInicio.Value)) And (IsNull(DtpFim.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelExtratoContrato.titu_dt_Vencimento >= " & FunNuloData(DtpInicio.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODOVENCIMENTO = "Vencimento: A partir de " & Format$(DtpInicio.Value, "dd/mm/yy")
    ElseIf (IsNull(DtpInicio.Value)) And (Not IsNull(DtpFim.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "(ConsCAPRelExtratoContrato.titu_dt_Vencimento <= " & FunNuloData(DtpFim.Value, NomeSgbd) & ")"
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODOVENCIMENTO = "Vencimento: Até " & Format$(DtpFim.Value, "dd/mm/yy")
    Else
        XLT_FILTROPERIODOVENCIMENTO = "Vencimento: Geral"
    End If
    
     'Periodo de Pagamento
    If (Not IsNull(DtpInicioPagamento.Value)) And (Not IsNull(DtpFimPagamento.Value)) Then
       XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelExtratoContrato.titu_dt_Pagamento >= " & FunNuloData(DtpInicioPagamento.Value, NomeSgbd) & "" & _
        " AND ConsCAPRelExtratoContrato.titu_dt_Pagamento <= " & FunNuloData(DtpFimPagamento.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODOPAGAMENTO = "Pagamento: " & Format$(DtpInicioPagamento.Value, "dd/mm/yy") & " a " & Format$(DtpFimPagamento.Value, "dd/mm/yy")
    ElseIf (Not IsNull(DtpInicioPagamento.Value)) And (IsNull(DtpFimPagamento.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelExtratoContrato.titu_dt_Pagamento >= " & FunNuloData(DtpInicioPagamento.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODOPAGAMENTO = "Pagamento: A partir de " & Format$(DtpInicioPagamento.Value, "dd/mm/yy")
    ElseIf (IsNull(DtpInicioPagamento.Value)) And (Not IsNull(DtpFimPagamento.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelExtratoContrato.titu_dt_Pagamento <= " & FunNuloData(DtpFimPagamento.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODOPAGAMENTO = "Pagamento: Até " & Format$(DtpFimPagamento.Value, "dd/mm/yy")
    ElseIf (IsNull(DtpInicioPagamento.Value)) And (IsNull(DtpFimPagamento.Value)) Then
        XLT_FILTROPERIODOPAGAMENTO = "Pagamento: Geral"
    End If
    
         'Periodo de Deposito
    If (Not IsNull(DtpInicioDeposito.Value)) And (Not IsNull(DtpFimDeposito.Value)) Then
       XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelExtratoContrato.titu_dt_Deposito >= " & FunNuloData(DtpInicioDeposito.Value, NomeSgbd) & "" & _
        " AND ConsCAPRelExtratoContrato.titu_dt_Deposito <= " & FunNuloData(DtpFimDeposito.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODODEPOSITO = "Depósito: " & Format$(DtpInicioDeposito.Value, "dd/mm/yy") & " a " & Format$(DtpFimDeposito.Value, "dd/mm/yy")
    ElseIf (Not IsNull(DtpInicioDeposito.Value)) And (IsNull(DtpFimDeposito.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelExtratoContrato.titu_dt_Deposito >= " & FunNuloData(DtpInicioDeposito.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODODEPOSITO = "Depósito: A partir de " & Format$(DtpInicioDeposito.Value, "dd/mm/yy")
    ElseIf (IsNull(DtpInicioDeposito.Value)) And (Not IsNull(DtpFimDeposito.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelExtratoContrato.titu_dt_Deposito <= " & FunNuloData(DtpFimDeposito.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
         XLT_FILTROPERIODODEPOSITO = "Depósito: Até " & Format$(DtpFimDeposito.Value, "dd/mm/yy")
    Else
         XLT_FILTROPERIODODEPOSITO = "Depósito: Geral"
    End If
    
    'Contrato
    If MskContratoInicial.Text <> "    .    .  " And MskContratoFinal.Text <> "    .    .  " Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsCAPRelExtratoContrato.Contrato >= '" & MskContratoInicial.Text & "'" & _
      " AND ConsCAPRelExtratoContrato.Contrato <= '" & MskContratoFinal.Text & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTROCONTRATO = "Contrato: " & MskContratoInicial & " a " & MskContratoFinal
    ElseIf MskContratoInicial.Text <> "    .    " And MskContratoFinal.Text = "    .    " Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsCAPRelExtratoContrato.Contrato >= '" & MskContratoInicial.Text & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTROCONTRATO = "Contrato: A partir de " & MskContratoInicial
    ElseIf MskContratoInicial.Text = "    .    " And MskContratoFinal.Text <> "    .    " Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsCAPRelExtratoContrato.Contrato <= '" & MskContratoFinal.Text & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTROCONTRATO = "Contrato: Até " & MskContratoInicial
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
            " OR ConsCAPRelExtratoContrato.cont_tx_Status = '" & XLT_STATUS2 & "') "
            XGT_CONJUNCAO = " AND "
        Else
            XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "ConsCAPRelExtratoContrato.cont_tx_Status = '" & XLT_STATUS & "' "
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
     
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "ConsCAPRelExtratoContrato.titu_tx_NaturezaPlano = '" & XLT_NATUREZA & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTRONATUREZAPLANO = "Natureza do Plano: " & CboNaturezasPlanos.Text
    Else
        XGT_CONJUNCAO = " AND "
      XLT_FILTRONATUREZAPLANO = "Natureza do Plano: Geral"
    End If
    
    If CboTipoImovel.Text <> "" Then
        XLT_TIPOIMOVEL = CboTipoImovel.BoundText
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "ConsCAPRelExtratoContrato.tiim_cd_TipoImovel = '" & XLT_TIPOIMOVEL & "'"
        XGT_CONJUNCAO = " AND "
        XLT_FILTROTIPOIMOVEL = "Tipo Imóvel: " & CboTipoImovel.Text
    End If
    
    'Filtra Títulos Pagos
    XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "ConsCAPRelExtratoContrato.titu_dt_Pagamento is not null "
    
    'Filtra a empresa Ativa
    XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "ConsCAPRelExtratoContrato.empr_cd_Empresa = " & CStr(PCodEmpresa)
  
    'Ordenação
    'XGT_SELECAO = XGT_SELECAO + " ORDER BY titu_dt_Pagamento"
    XGT_SELECAO = XGT_SELECAO + " ORDER BY Titulo"
    
    'Abre os recordsets e verifica se eles possuem dados
    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute(XLT_SQL & XGT_SELECAO)
    If XGO_RSRELATORIO.EOF Then
        MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
        FrmRelTitulosLiquidados.MousePointer = vbDefault
        Exit Sub
    End If
    Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
         
    Select Case XOrigem
      Case "Analitico"
        FrmVisRelatorios.Caption = "Relatório Títulos Liquidados - Analítico"
        If ChkEmpreendimentos.Value = 1 Then
          'Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelTitulosLiquidadosAnaliticoAgrupados.rpt")
          If ChkApresentarValorInicial.Value = 0 Then
            Set XGR_RELATORIO = RptTitulosLiquidadosAnaliticoAgrupados
          Else
            Set XGR_RELATORIO = RptTitulosLiquidadosAnaliticoAgrupadosValorInicial
          End If
        Else
          'Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelTitulosLiquidadosAnalitico.rpt")
          If Me.chkJurosContratuais.Value = 1 Then
            Set XGR_RELATORIO = RptTitulosLiquidadosAnaliticoJurosC
          Else
            Set XGR_RELATORIO = RptTitulosLiquidadosAnalitico
          End If
          'Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Teste.rpt")
        End If
      Case "Sintetico"
        If ChkCorrecaoMonetaria.Value = 1 Then 'Destacar Correção Monetária
            FrmVisRelatorios.Caption = "Relatório Títulos Liquidados - Sintético - Com Correção Monetária Destacada"
            Set XGR_RELATORIO = RptTitulosLiquidadosSinteticoCMonetariaDestacada
        ElseIf ChkApresentarValorInicial.Value = 1 Then
            FrmVisRelatorios.Caption = "Relatório Títulos Liquidados - Sintético - Vl Original sem Correção"
            Set XGR_RELATORIO = RptTitulosLiquidadosSinteticoValorOriginal
        Else
            FrmVisRelatorios.Caption = "Relatório Títulos Liquidados - Sintético"
            'Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelTitulosLiquidadosSintetico.rpt")
            Set XGR_RELATORIO = RptTitulosLiquidadosSintetico
        End If
      Case "PorNatureza"
 '       Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelTitulosLiquidadosPorNatureza.rpt")
'        Set XGR_RELATORIO = RptTitulosLiquidadosPorNatureza
    End Select
    
    XGR_RELATORIO.DiscardSavedData
    
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
  
    'Ativa os Relatórios
    With XGR_RELATORIO.FormulaFields
        .GetItemByName("FiltroEmpreendimento").Text = Chr(34) & XLT_FILTROCONTRATO & Chr(34)
        .GetItemByName("FiltroPeriodoVencimento").Text = Chr(34) & XLT_FILTROPERIODOVENCIMENTO & Chr(34)
        .GetItemByName("FiltroPeriodoPagamento").Text = Chr(34) & XLT_FILTROPERIODOPAGAMENTO & Chr(34)
        .GetItemByName("FiltroPeriodoDeposito").Text = Chr(34) & XLT_FILTROPERIODODEPOSITO & Chr(34)
        .GetItemByName("FiltroStatus").Text = Chr(34) & XLT_FILTROSTATUS & Chr(34)
        .GetItemByName("FiltroTipoPlano").Text = Chr(34) & XLT_FILTROTIPOPLANO & Chr(34)
        .GetItemByName("FiltroNaturezaPlano").Text = Chr(34) & XLT_FILTRONATUREZAPLANO & Chr(34)
        .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
        
        If ChkCorrecaoMonetaria.Value = 1 Then
            .GetItemByName("Indexador").Text = Chr(34) & "1" & Chr(34)
        Else
            .GetItemByName("Indexador").Text = Chr(34) & "0" & Chr(34)
        End If
    End With
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
    
    If XOrigem = "Analitico" Then
      FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSRELATORIO, "TotalIndexador"
    End If
    
    FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
    FrmRelTitulosLiquidados.MousePointer = vbDefault
End Sub

Private Sub CmdLimparNaturezaPlano_Click()
  
  CboNaturezasPlanos.ListIndex = -1
  
End Sub

Private Sub CmdLimparTipoPlano_Click()
  
  CboTiposPlanos.BoundText = ""
  
End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelTitulosLiquidados
End Sub

Private Sub Form_Load()
  
  subManutencaoJanelasAtivas "I", "FrmRelTitulosLiquidados"
  subConectarControleDadosNV DatTipoPlano, "SELECT * FROM TiposPlanos ORDER BY tipl_tx_Descricao", Estatico
  subConectarControleDadosNV DatTipoImovel, "SELECT * FROM TiposImoveis ORDER BY tiim_tx_Descricao", Estatico
  
  Select Case XOrigem
    Case "Analitico"
      Me.Caption = "Relatório Titulos Liquidados - Analítico"
      ChkCorrecaoMonetaria.Visible = False
      
    Case "Sintetico"
      Me.Caption = "Relatório Titulos Liquidados - Sintético"
      FraAgrupados.Visible = False
      ChkCorrecaoMonetaria.Visible = True
      ChkCorrecaoMonetaria.Top = 1620
      ChkCorrecaoMonetaria.Left = 90
      
  End Select
  
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
    If Len(Trim(MskContratoInicial.Text)) < 12 Then
      MsgBox "O código inicial do contrato deve possuir 12 dígitos, ou deve estar vazio.", vbCritical, "ATENÇÃO !"
      MskContratoInicial.SetFocus
    End If
  End If
  
End Sub

Private Sub MskContratoFinal_LostFocus()
   
  If MskContratoFinal.Text <> "    .    .  " Then
    If Len(Trim(MskContratoFinal.Text)) < 12 Then
      MsgBox "O código final do imóvel deve possuir 12 dígitos, ou deve estar vazio", vbCritical, "ATENÇÃO !"
      MskContratoFinal.SetFocus
    End If
  End If
  
End Sub
