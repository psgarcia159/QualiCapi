VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRelTitulosVencidos 
   Caption         =   "Relatório Títulos Vencidos"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   Icon            =   "FrmRelTitulosVencidos.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4110
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   3555
      Left            =   4620
      TabIndex        =   34
      Top             =   0
      Width           =   4695
      Begin VB.Frame FraProrrogados 
         Height          =   555
         Left            =   60
         TabIndex        =   41
         Top             =   2880
         Width           =   4575
         Begin VB.CheckBox ChkProrrogados 
            Alignment       =   1  'Right Justify
            Caption         =   "Considerar Títulos Prorrogados?"
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
            TabIndex        =   17
            Top             =   180
            Value           =   1  'Checked
            Width           =   3075
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
         Height          =   795
         Left            =   60
         TabIndex        =   37
         Top             =   720
         Width           =   4575
         Begin VB.ComboBox CboNaturezasPlanos 
            Height          =   315
            ItemData        =   "FrmRelTitulosVencidos.frx":2AFA
            Left            =   1740
            List            =   "FrmRelTitulosVencidos.frx":2B10
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   300
            Width           =   2415
         End
         Begin Threed.SSCommand CmdLimparNaturezaPlano 
            Height          =   315
            Left            =   4185
            TabIndex        =   14
            Top             =   300
            Width           =   315
            _Version        =   65536
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "FrmRelTitulosVencidos.frx":2B4B
            Picture         =   "FrmRelTitulosVencidos.frx":2B67
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
            TabIndex        =   38
            Top             =   360
            Width           =   1680
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
         Height          =   555
         Left            =   60
         TabIndex        =   35
         Top             =   120
         Width           =   4575
         Begin Threed.SSCommand CmdLimparTipoPlano 
            Height          =   315
            Left            =   4200
            TabIndex        =   12
            Top             =   180
            Width           =   315
            _Version        =   65536
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "FrmRelTitulosVencidos.frx":2C79
            Picture         =   "FrmRelTitulosVencidos.frx":2C95
         End
         Begin MSDataListLib.DataCombo CboTiposPlanos 
            Bindings        =   "FrmRelTitulosVencidos.frx":2DA7
            Height          =   315
            Left            =   1740
            TabIndex        =   11
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
            Left            =   420
            TabIndex        =   36
            Top             =   240
            Width           =   1260
         End
      End
      Begin VB.Frame FraTitulosPagos 
         Height          =   675
         Left            =   60
         TabIndex        =   40
         Top             =   1500
         Width           =   4575
         Begin VB.CheckBox ChkTitulosPagos 
            Alignment       =   1  'Right Justify
            Caption         =   "Considerar Títulos Pagos?"
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
            Left            =   60
            TabIndex        =   15
            Top             =   240
            Width           =   2955
         End
      End
      Begin VB.Frame FraAgrupados 
         Height          =   675
         Left            =   60
         TabIndex        =   39
         Top             =   2160
         Width           =   4575
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
            Height          =   315
            Left            =   60
            TabIndex        =   16
            Top             =   240
            Value           =   1  'Checked
            Width           =   2955
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3555
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   4575
      Begin VB.Frame FraDataBase 
         Height          =   555
         Left            =   60
         TabIndex        =   32
         Top             =   2880
         Width           =   4455
         Begin MSComCtl2.DTPicker DtpDtBase 
            Height          =   315
            Left            =   1140
            TabIndex        =   10
            Top             =   180
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MM/yy"
            Format          =   48889859
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
            TabIndex        =   33
            Top             =   240
            Width           =   1050
         End
      End
      Begin VB.Frame FraImovel 
         Height          =   555
         Left            =   60
         TabIndex        =   29
         Top             =   120
         Width           =   4455
         Begin MSMask.MaskEdBox MskImovelInicial 
            Height          =   315
            Left            =   960
            TabIndex        =   0
            Top             =   180
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   9
            Mask            =   "AAAA.AAAA"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MskImovelFinal 
            Height          =   315
            Left            =   2460
            TabIndex        =   1
            Top             =   180
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   9
            Mask            =   "AAAA.AAAA"
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
            Left            =   2160
            TabIndex        =   31
            Top             =   240
            Width           =   195
         End
         Begin VB.Label LblImovel 
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
            Height          =   285
            Left            =   240
            TabIndex        =   30
            Top             =   240
            Width           =   705
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
         TabIndex        =   26
         Top             =   2160
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
            TabIndex        =   7
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
            TabIndex        =   8
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
            TabIndex        =   9
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
         Height          =   765
         Left            =   60
         TabIndex        =   23
         Top             =   720
         Width           =   4470
         Begin MSComCtl2.DTPicker DtpInicio 
            Height          =   315
            Left            =   660
            TabIndex        =   2
            Top             =   285
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   48889859
            CurrentDate     =   37677.572037037
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
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   48889859
            CurrentDate     =   37614
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
            TabIndex        =   25
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
            TabIndex        =   24
            Top             =   360
            Width           =   495
         End
      End
      Begin Threed.SSFrame FraStatus 
         Height          =   660
         Left            =   60
         TabIndex        =   27
         Top             =   1500
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
            TabIndex        =   5
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
            TabIndex        =   4
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
            TabIndex        =   6
            Top             =   300
            Width           =   1230
         End
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir..."
      Height          =   345
      Left            =   6780
      TabIndex        =   20
      Top             =   3660
      Width           =   1035
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   8160
      TabIndex        =   21
      Top             =   3660
      Width           =   1035
   End
   Begin MSDataListLib.DataCombo CboObs 
      Bindings        =   "FrmRelTitulosVencidos.frx":2DC2
      Height          =   315
      Left            =   1260
      TabIndex        =   18
      Top             =   3660
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
      Left            =   6120
      TabIndex        =   19
      Top             =   3660
      Width           =   360
      _Version        =   65536
      _ExtentX        =   635
      _ExtentY        =   556
      _StockProps     =   78
      MouseIcon       =   "FrmRelTitulosVencidos.frx":2DD7
      Picture         =   "FrmRelTitulosVencidos.frx":3229
   End
   Begin MSAdodcLib.Adodc DatObs 
      Height          =   330
      Left            =   3660
      Top             =   3660
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
      Left            =   60
      TabIndex        =   28
      Top             =   3720
      Width           =   1140
   End
End
Attribute VB_Name = "FrmRelTitulosVencidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
    
    Dim XLT_FILTROPERIODO As String 'Armazenar filtro do período escolhido
    Dim XLT_FILTROIMOVEL As String
    Dim XLT_FILTROSTATUS As String
    Dim XLT_FILTROTIPOPLANO As String
    Dim XLT_FILTRONATUREZAPLANO As String
    Dim XLT_FILTROEMPREENDIMENTOS As String
    Dim XLT_FILTROTITULOSPAGOS As String
    Dim XLT_FILTROMULTA As String
    Dim XLT_FILTROJUROS As String
    Dim XLO_TOTALINDEXADOR As Object
    Dim XLT_FILTROPRORROGADOS As String
    Dim XLT_FILTROOBSERVACAO As String
    
    Dim XLT_STATUS As String
    Dim XLT_STATUS2 As String
    Dim XLT_NATUREZA As String 'Natureza do Plano
    Dim XLT_INDEXADOR As String
    
    Dim XLT_SQL As String
    
    If DtpDtBase.Value < DtpFim.Value Then
      MsgBox "A data base deve ser maior ou igual ao período final de vencimento", vbCritical, "ATENÇÂO"
      Exit Sub
    End If
    
    FrmRelTitulosVencidos.MousePointer = vbHourglass
    
    XGT_SELECAO = ""
    XGT_CONJUNCAO = ""
    
    If NomeSgbd = "Access" Then
    
      XLT_SQL = "" & _
      "SELECT ConsCAPRelPrevisaoRecebimento.*," & _
      "IIF(Indexacao1='M'" & _
        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & "))" & _
      " as Cotacao1," & _
      "IIF(Indexacao2='M'" & _
        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & "))" & _
      "  as Cotacao2,"
      
'      XLT_SQL = XLT_SQL & _
'      "IIF(Indexacao1='M'" & _
'        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")))" & _
'      " as Cotacao1_MesAnterior," & _
'      "IIF(Indexacao2='M'" & _
'        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")))" & _
'      " as Cotacao2_MesAnterior" & _
'      " From ConsCAPRelPrevisaoRecebimento WHERE "
      
      XLT_SQL = XLT_SQL & _
      "IIF(Indexacao1='M'" & _
        " ,IIF (cont_tx_CorrecProrata = 'A',(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
        " ,IIF (cont_tx_CorrecProrata = 'P',(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")))))" & _
      " as Cotacao1_MesAnterior," & _
      "IIF(Indexacao2='M'" & _
        " ,IIF (cont_tx_CorrecProrata = 'A',(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
        " ,IIF (cont_tx_CorrecProrata = 'P',(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")))))" & _
      " as Cotacao2_MesAnterior" & _
      " From ConsCAPRelPrevisaoRecebimento WHERE "
    
    Else
      XLT_SQL = "" & _
      "SELECT ConsCAPRelPrevisaoRecebimento.*," & _
      "CASE Indexacao1   when 'M' then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
        " when 'D' then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & ")" & _
      " end as Cotacao1," & _
      "CASE  Indexacao2 when 'M' then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
        " when 'D'  then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & ")" & _
      " end as Cotacao2,"
      
'      XLT_SQL = XLT_SQL & _
'      "CASE  Indexacao1 when 'M' then" & _
'        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
'      " end as Cotacao1_MesAnterior," & _
'      "CASE Indexacao2 when 'M' then" & _
'        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
'      " end As Cotacao2_MesAnterior" & _
'      " From ConsCAPRelPrevisaoRecebimento WHERE "

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
      " end As Cotacao2_MesAnterior" & _
      " From ConsCAPRelPrevisaoRecebimento WHERE "
      
    End If
    
    'Período
    If ChkProrrogados.Value = 1 Then
                 
      'Período de Prorrogacao
       If (Not IsNull(DtpInicio.Value)) And (Not IsNull(DtpFim.Value)) Then
          XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
           "ConsCAPRelPrevisaoRecebimento.titu_dt_Prorrogacao >= " & FunNuloData(DtpInicio.Value, NomeSgbd) & "" & _
           " AND ConsCAPRelPrevisaoRecebimento.titu_dt_Prorrogacao <= " & FunNuloData(DtpFim.Value, NomeSgbd) & ""
           XGT_CONJUNCAO = " AND "
           XLT_FILTROPERIODO = "Vencimento: " & Format$(DtpInicio.Value, "dd/mm/yy") & " a " & Format$(DtpFim.Value, "dd/mm/yy")
       ElseIf (Not IsNull(DtpInicio.Value)) And (IsNull(DtpFim.Value)) Then
           XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
           "ConsCAPRelPrevisaoRecebimento.titu_dt_Prorrogacao >= " & FunNuloData(DtpInicio.Value, NomeSgbd) & ""
           XGT_CONJUNCAO = " AND "
           XLT_FILTROPERIODO = "Vencimento: A partir de " & Format$(DtpInicio.Value, "dd/mm/yy")
       ElseIf (IsNull(DtpInicio.Value)) And (Not IsNull(DtpFim.Value)) Then
           XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
           "(ConsCAPRelPrevisaoRecebimento.titu_dt_Prorrogacao <= " & FunNuloData(DtpFim.Value, NomeSgbd) & ")"
           XGT_CONJUNCAO = " AND "
           XLT_FILTROPERIODO = "Vencimento: Até " & Format$(DtpFim.Value, "dd/mm/yy")
       Else
           XLT_FILTROPERIODO = "Vencimento: Geral"
       End If
      
      XLT_FILTROPRORROGADOS = "1"
    Else
       
      'Período de Vencimento
      If (Not IsNull(DtpInicio.Value)) And (Not IsNull(DtpFim.Value)) Then
         XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
          "ConsCAPRelPrevisaoRecebimento.titu_dt_Vencimento >= " & FunNuloData(DtpInicio.Value, NomeSgbd) & "" & _
          " AND ConsCAPRelPrevisaoRecebimento.titu_dt_Vencimento <= " & FunNuloData(DtpFim.Value, NomeSgbd) & ""
          XGT_CONJUNCAO = " AND "
          XLT_FILTROPERIODO = "Vencimento: " & Format$(DtpInicio.Value, "dd/mm/yy") & " a " & Format$(DtpFim.Value, "dd/mm/yy")
      ElseIf (Not IsNull(DtpInicio.Value)) And (IsNull(DtpFim.Value)) Then
          XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
          "ConsCAPRelPrevisaoRecebimento.titu_dt_Vencimento >= " & FunNuloData(DtpInicio.Value, NomeSgbd) & ""
          XGT_CONJUNCAO = " AND "
          XLT_FILTROPERIODO = "Vencimento: A partir de " & Format$(DtpInicio.Value, "dd/mm/yy")
      ElseIf (IsNull(DtpInicio.Value)) And (Not IsNull(DtpFim.Value)) Then
          XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
          "(ConsCAPRelPrevisaoRecebimento.titu_dt_Vencimento <= " & FunNuloData(DtpFim.Value, NomeSgbd) & ")"
          XGT_CONJUNCAO = " AND "
          XLT_FILTROPERIODO = "Vencimento: Até " & Format$(DtpFim.Value, "dd/mm/yy")
      Else
          XLT_FILTROPERIODO = "Vencimento: Geral"
      End If
      
       XLT_FILTROPRORROGADOS = "0"
    End If
    
    'Imóvel
    If MskImovelInicial.Text <> "    .    " And MskImovelFinal.Text <> "    .    " Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsCAPRelPrevisaoRecebimento.Imovel >= '" & MskImovelInicial.Text & "'" & _
      " AND ConsCAPRelPrevisaoRecebimento.Imovel <= '" & MskImovelFinal.Text & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTROIMOVEL = "Imóvel: " & MskImovelInicial & " a " & MskImovelFinal
    ElseIf MskImovelInicial.Text <> "    .    " And MskImovelFinal.Text = "    .    " Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsCAPRelPrevisaoRecebimento.Imovel >= '" & MskImovelInicial.Text & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTROIMOVEL = "Imóvel: A partir de " & MskImovelInicial
    ElseIf MskImovelInicial.Text = "    .    " And MskImovelFinal.Text <> "    .    " Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsCAPRelPrevisaoRecebimento.Imovel <= '" & MskImovelFinal.Text & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTROIMOVEL = "Imóvel: Até " & MskImovelInicial
    Else
      XLT_FILTROIMOVEL = "Imóvel: Geral"
    End If
    
    'Status
    If ChkAtivo.Value = ChkInativo.Value And ChkAtivo.Value = ChkCedido.Value Then
       XLT_FILTROSTATUS = "Status: Geral"
       XLT_STATUS = "T" 'Todos
    Else
      If ChkAtivo.Value = 1 Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
          "(ConsCAPRelPrevisaoRecebimento.cont_dt_Venda <= " & FunNuloData(DtpDtBase.Value, NomeSgbd) & _
          " AND (ConsCAPRelPrevisaoRecebimento.cont_dt_Distrato is null OR " & _
          " (ConsCAPRelPrevisaoRecebimento.cont_dt_Distrato > " & FunNuloData(DtpDtBase.Value, NomeSgbd) & " AND " & _
          " ConsCAPRelPrevisaoRecebimento.cont_tx_Status<>'C')) "
        
        If ChkInativo.Value = 1 Then  'Ativo e Inativo
          XGT_SELECAO = XGT_SELECAO & _
            " OR (ConsCAPRelPrevisaoRecebimento.cont_tx_Status = 'I'" & _
            " AND ConsCAPRelPrevisaoRecebimento.cont_dt_Distrato<=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & "))"
            
          XLT_STATUS = "A"
          XLT_STATUS2 = "I"
          XLT_FILTROSTATUS = "Status: Ativo e Inativo"
        
        ElseIf ChkCedido.Value = 1 Then  'Ativo e Reservado
          XGT_SELECAO = XGT_SELECAO & _
            " OR (ConsCAPRelPrevisaoRecebimento.cont_tx_Status = 'C'" & _
            " AND ConsCAPRelPrevisaoRecebimento.cont_dt_Distrato<=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & "))"
            
          XLT_STATUS = "A"
          XLT_STATUS2 = "C"
          XLT_FILTROSTATUS = "Status: Ativo e Cedido"
          
        Else 'Ativo
          XGT_SELECAO = XGT_SELECAO & ")" 'Fecha o parêntese
          XLT_STATUS = "A"
          XLT_FILTROSTATUS = "Status: Ativo"
        End If
    XGT_CONJUNCAO = " AND "
        
      ElseIf ChkInativo.Value = 1 Then
        If ChkCedido.Value = 1 Then   'Inativo e Cedido
          XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
            " ((ConsCAPRelPrevisaoRecebimento.cont_tx_Status = 'I'" & _
            " OR ConsCAPRelPrevisaoRecebimento.cont_tx_Status = 'C')" & _
            " AND ConsCAPRelPrevisaoRecebimento.cont_dt_Distrato<=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & ")"
            
           XLT_STATUS = "I"
           XLT_STATUS2 = "C"
           XLT_FILTROSTATUS = "Status: Inativo e Cedido"
        Else 'Inativo
          XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
            " (ConsCAPRelPrevisaoRecebimento.cont_tx_Status = 'I'" & _
            " AND ConsCAPRelPrevisaoRecebimento.cont_dt_Distrato<=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & ")"
            
          XLT_STATUS = "I"
          XLT_FILTROSTATUS = "Status: Inativo"
        End If
    
      ElseIf ChkCedido.Value = 1 Then 'Cedido
         XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
           " (ConsCAPRelPrevisaoRecebimento.cont_tx_Status = 'C'" & _
           " AND ConsCAPRelPrevisaoRecebimento.cont_dt_Distrato<=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & ")"
            
         XLT_STATUS = "C"
         XLT_FILTROSTATUS = "Status: Cedido"
      End If
    End If
    
    'Tipo do Plano
    If CboTiposPlanos.BoundText <> "" Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO + "ConsCAPRelPrevisaoRecebimento.tipl_cd_TipoPlano = " & CboTiposPlanos.BoundText & ""
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
     
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO + "ConsCAPRelPrevisaoRecebimento.titu_tx_NaturezaPlano = '" & XLT_NATUREZA & "'"
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
    
    'Titulos Pagos
    If ChkTitulosPagos.Value = 1 Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      " ((ConsCAPRelPrevisaoRecebimento.titu_dt_Pagamento is Null AND " & _
      " ConsCAPRelPrevisaoRecebimento.tipl_tx_Sigla <> 'R')"
      
      XGT_SELECAO = XGT_SELECAO & _
      " OR(ConsCAPRelPrevisaoRecebimento.titu_dt_Pagamento is Null AND " & _
      " ConsCAPRelPrevisaoRecebimento.tipl_tx_Sigla = 'R' AND " & _
      " ConsCAPRelPrevisaoRecebimento.titu_dt_Base < " & FunNuloData(DtpDtBase.Value, NomeSgbd) & ") "
      
      XGT_SELECAO = XGT_SELECAO & _
      " OR (ConsCAPRelPrevisaoRecebimento.titu_dt_Pagamento is not Null AND " & _
      " ConsCAPRelPrevisaoRecebimento.tipl_tx_Sigla = 'R' AND " & _
      " ConsCAPRelPrevisaoRecebimento.titu_dt_Base < " & FunNuloData(DtpDtBase.Value, NomeSgbd) & " AND " & _
      " ConsCAPRelPrevisaoRecebimento.titu_dt_Pagamento >= " & FunNuloData(DtpDtBase.Value, NomeSgbd) & ")"
      
      XGT_SELECAO = XGT_SELECAO & _
      " OR (ConsCAPRelPrevisaoRecebimento.titu_dt_Pagamento is not Null AND " & _
      " ConsCAPRelPrevisaoRecebimento.tipl_tx_Sigla <> 'R' AND " & _
      " ConsCAPRelPrevisaoRecebimento.titu_dt_Pagamento >= " & FunNuloData(DtpDtBase.Value, NomeSgbd) & ")) "
      
      XGT_CONJUNCAO = " AND "
      XLT_FILTROTITULOSPAGOS = "Considerando os Títulos Pagos"
    Else
    
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsCAPRelPrevisaoRecebimento.titu_dt_Pagamento is Null"
      XGT_CONJUNCAO = " AND "
      XLT_FILTROTITULOSPAGOS = "Não Considerando os Títulos Pagos"
    End If
      
    'Vencidos
    If ChkProrrogados.Value = 1 Then
     XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "titu_dt_Prorrogacao < " & FunNuloData(DtpFim.Value, NomeSgbd)
      XGT_CONJUNCAO = " AND "
    Else
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "titu_dt_Vencimento < " & FunNuloData(DtpFim.Value, NomeSgbd)
      XGT_CONJUNCAO = " AND "
    End If
    
    'Observacao
    If CboObs.BoundText <> "" Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsCAPRelPrevisaoRecebimento.obse_cd_Observacao=" & CboObs.BoundText & " "
      XGT_CONJUNCAO = " AND "
      XLT_FILTROOBSERVACAO = "Observação: " & CboObs.Text
    Else
      XLT_FILTROOBSERVACAO = "Observação:Geral"
    End If
    
    'Filtra a empresa Ativa
    XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + "ConsCAPRelPrevisaoRecebimento.empr_cd_Empresa = " & CStr(PCodEmpresa)
  
    'Ordena
    If ChkProrrogados.Value = 1 Then
      XGT_SELECAO = XGT_SELECAO + " ORDER BY titu_dt_Prorrogacao," & _
       " Titulo"
    Else
      XGT_SELECAO = XGT_SELECAO + " ORDER BY titu_dt_Vencimento," & _
      " Titulo"
    End If
  
    'Abre os recordsets e verifica se eles possuem dados
     Set XGO_RSRELATORIO = ConexaoRelatorio.Execute(XLT_SQL & XGT_SELECAO)
    If XGO_RSRELATORIO.EOF Then
        MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
        FrmRelTitulosVencidos.MousePointer = vbDefault
        Exit Sub
    End If
    Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
    
    FrmVisRelatorios.Caption = "Relatório Títulos Vencidos"
    If ChkEmpreendimentos.Value = 1 Then
      'Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelTitulosVencidosAgrupados.rpt")
      Set XGR_RELATORIO = RptTitulosVencidosAgrupados
    Else
      'Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelTitulosVencidos.rpt")
      Set XGR_RELATORIO = RptTitulosVencidos
    End If
    XGR_RELATORIO.DiscardSavedData
  
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
  
    'Ativa os Relatórios
    With XGR_RELATORIO.FormulaFields
        .GetItemByName("FiltroPeriodo").Text = Chr(34) & XLT_FILTROPERIODO & Chr(34)
        .GetItemByName("FiltroImovel").Text = Chr(34) & XLT_FILTROIMOVEL & Chr(34)
        .GetItemByName("FiltroStatus").Text = Chr(34) & XLT_FILTROSTATUS & Chr(34)
        .GetItemByName("FiltroTipoPlano").Text = Chr(34) & XLT_FILTROTIPOPLANO & Chr(34)
        .GetItemByName("FiltroNaturezaPlano").Text = Chr(34) & XLT_FILTRONATUREZAPLANO & Chr(34)
        .GetItemByName("FiltroDataBase").Text = Chr(34) & Format(DtpDtBase.Value, "dd/mm/yy") & Chr(34)
        .GetItemByName("Indexador").Text = Chr(34) & XLT_INDEXADOR & Chr(34)
        '.GetItemByName("FiltroEmpreendimentos").Text = Chr(34) & XLT_FILTROEMPREENDIMENTOS & Chr(34)
        .GetItemByName("FiltroProrrogados").Text = Chr(34) & XLT_FILTROPRORROGADOS & Chr(34)
        .GetItemByName("FiltroTitulosPagos").Text = Chr(34) & XLT_FILTROTITULOSPAGOS & Chr(34)
        .GetItemByName("FiltroObservacao").Text = Chr(34) & XLT_FILTROOBSERVACAO & Chr(34)
        .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
    End With
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
    
    FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
    FrmRelTitulosVencidos.MousePointer = vbDefault
End Sub

Private Sub CmdLimparNaturezaPlano_Click()
  
  CboNaturezasPlanos.ListIndex = -1
  
End Sub

Private Sub CmdLimparTipoPlano_Click()
  
  CboTiposPlanos.BoundText = ""
  
End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelTitulosVencidos
End Sub

Private Sub Form_Load()
  
  subManutencaoJanelasAtivas "I", "FrmRelTitulosVencidos"
  subConectarControleDadosNV DatTipoPlano, "SELECT * FROM TiposPlanos ORDER BY tipl_tx_Descricao", Estatico
  subConectarControleDadosNV DatObs, "SELECT * FROM Observacoes ORDER BY obse_tx_observacao", Estatico
  
  DtpDtBase.Value = Date
  DtpFim.Value = Date
  
End Sub

Private Sub MskImovelInicial_GotFocus()
  
  Call subSelecionaMSK(MskImovelInicial)
  
End Sub

Private Sub MskImovelFinal_GotFocus()
  
  Call subSelecionaMSK(MskImovelFinal)
  
End Sub

Private Sub MskImovelInicial_KeyPress(KeyAscii As Integer)
  If KeyAscii >= 97 And KeyAscii <= 122 Then
    KeyAscii = KeyAscii - 32
  End If
End Sub

Private Sub MskImovelFinal_KeyPress(KeyAscii As Integer)
  If KeyAscii >= 97 And KeyAscii <= 122 Then
    KeyAscii = KeyAscii - 32
  End If
End Sub

Private Sub MskImovelInicial_LostFocus()
   
  If MskImovelInicial.Text <> "    .    " Then
    If Len(Trim(MskImovelInicial.Text)) < 9 Then
      MsgBox "O código inicial do imóvel deve possuir 8 dígitos, ou deve estar vazio.", vbCritical, "ATENÇÃO !"
      MskImovelInicial.SetFocus
    End If
  End If
  
End Sub

Private Sub MskImovelFinal_LostFocus()
   
  If MskImovelFinal.Text <> "    .    " Then
    If Len(Trim(MskImovelFinal.Text)) < 9 Then
      MsgBox "O código final do imóvel deve possuir 8 dígitos, ou deve estar vazio", vbCritical, "ATENÇÃO !"
      MskImovelFinal.SetFocus
    End If
  End If
  
End Sub
