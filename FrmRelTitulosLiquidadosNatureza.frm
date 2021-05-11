VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRelTitulosLiquidadosNatureza 
   Caption         =   "Títulos Liquidados Por Natureza"
   ClientHeight    =   3480
   ClientLeft      =   945
   ClientTop       =   2190
   ClientWidth     =   8640
   Icon            =   "FrmRelTitulosLiquidadosNatureza.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3480
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraContrato 
      Height          =   495
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   8625
      Begin MSMask.MaskEdBox MskContratoInicial 
         Height          =   315
         Left            =   3180
         TabIndex        =   31
         Top             =   120
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
         Left            =   4740
         TabIndex        =   32
         Top             =   120
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
         Left            =   4560
         TabIndex        =   34
         Top             =   180
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
         Left            =   2280
         TabIndex        =   33
         Top             =   180
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2595
      Left            =   4620
      TabIndex        =   19
      Top             =   480
      Width           =   4005
      Begin VB.Frame FraAcrescimos 
         Caption         =   "Apresentar títulos pagos com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   60
         TabIndex        =   23
         Top             =   1015
         Width           =   3885
         Begin VB.CheckBox CheDesagio 
            Alignment       =   1  'Right Justify
            Caption         =   "Deságio?"
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
            Left            =   1980
            TabIndex        =   12
            Top             =   1020
            Value           =   1  'Checked
            Width           =   1275
         End
         Begin VB.CheckBox ChkOutros 
            Alignment       =   1  'Right Justify
            Caption         =   "Outros?"
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
            Left            =   360
            TabIndex        =   9
            Top             =   1020
            Value           =   1  'Checked
            Width           =   1035
         End
         Begin VB.CheckBox ChkJuros 
            Alignment       =   1  'Right Justify
            Caption         =   "Juros?"
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
            Left            =   360
            TabIndex        =   8
            Top             =   660
            Value           =   1  'Checked
            Width           =   1035
         End
         Begin VB.CheckBox ChkMulta 
            Alignment       =   1  'Right Justify
            Caption         =   "Multa?"
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
            Left            =   360
            TabIndex        =   7
            Top             =   300
            Value           =   1  'Checked
            Width           =   1035
         End
         Begin VB.CheckBox ChkSeguro 
            Alignment       =   1  'Right Justify
            Caption         =   "Seguro?"
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
            Left            =   1980
            TabIndex        =   10
            Top             =   300
            Value           =   1  'Checked
            Width           =   1275
         End
         Begin VB.CheckBox ChkDesconto 
            Alignment       =   1  'Right Justify
            Caption         =   "Desconto?"
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
            Left            =   1980
            TabIndex        =   11
            Top             =   660
            Value           =   1  'Checked
            Width           =   1275
         End
      End
      Begin Threed.SSFrame FraStatus 
         Height          =   660
         Left            =   60
         TabIndex        =   24
         Top             =   120
         Width           =   3885
         _Version        =   65536
         _ExtentX        =   6853
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
            Left            =   2760
            TabIndex        =   6
            Top             =   300
            Width           =   975
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
            Left            =   240
            TabIndex        =   4
            Top             =   300
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
            Left            =   1440
            TabIndex        =   5
            Top             =   300
            Width           =   1080
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2595
      Left            =   0
      TabIndex        =   15
      Top             =   480
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
         Height          =   675
         Left            =   60
         TabIndex        =   25
         Top             =   1800
         Width           =   4470
         Begin MSComCtl2.DTPicker DtpInicioDeposito 
            Height          =   315
            Left            =   660
            TabIndex        =   26
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   69599235
            CurrentDate     =   37617.6519560185
         End
         Begin MSComCtl2.DTPicker DtpFimDeposito 
            Height          =   315
            Left            =   2760
            TabIndex        =   27
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   69599235
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
            TabIndex        =   29
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
            TabIndex        =   28
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
         TabIndex        =   20
         Top             =   960
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
            Format          =   69599235
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
            Format          =   69599235
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
            TabIndex        =   22
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
            TabIndex        =   21
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
         TabIndex        =   16
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
            Format          =   69599235
            CurrentDate     =   37690.4234490741
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
            Format          =   69599235
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
            TabIndex        =   18
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
            TabIndex        =   17
            Top             =   360
            Width           =   495
         End
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir..."
      Height          =   345
      Left            =   6120
      TabIndex        =   13
      Top             =   3120
      Width           =   1035
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   7440
      TabIndex        =   14
      Top             =   3120
      Width           =   1035
   End
End
Attribute VB_Name = "FrmRelTitulosLiquidadosNatureza"
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
    Dim XLT_FILTROSTATUS As String
    Dim XLT_FILTRONATUREZAPLANO As String
    Dim XLT_FILTROCONTRATO As String
    Dim XLT_FILTROMULTA As String
    Dim XLT_FILTROJUROS As String
    Dim XLT_FILTROOUTROS As String
    Dim XLT_FILTROSEGURO As String
    Dim XLT_FILTRODESCONTO As String
    Dim XLT_FILTRODESAGIO As String
    
    Dim XLT_FILTROPERIODOPAGAMENTO As String
    
    Dim XLT_STATUS As String
    Dim XLT_STATUS2 As String
    Dim XLT_INDEXADOR As String
    Dim XLT_SQL As String
    
    FrmRelTitulosLiquidadosNatureza.MousePointer = vbHourglass
    
    XGT_SELECAO = ""
    XGT_CONJUNCAO = ""
    
    If NomeSgbd = "Access" Then
      
      XLT_SQL = "" & _
      "SELECT ConsCAPRelExtratoContrato.*," & _
        "IIF(Indexacao3='M'" & _
        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=format(titu_dt_BasePagto,'\01/mm/yy'))" & _
        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=titu_dt_BasePagto))" & _
      " as Cotacao3,"
        
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
    Else
        XLT_FILTROPERIODOPAGAMENTO = "Pagamento: Geral"
    End If
    
    
            'Periodo de Deposito
    If (Not IsNull(DtpInicioDeposito.Value)) And (Not IsNull(DtpFimDeposito.Value)) Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelExtratoContrato.titu_dt_Deposito >= " & FunNuloData(DtpInicioDeposito.Value, NomeSgbd) & "" & _
        " AND ConsCAPRelExtratoContrato.titu_dt_Deposito <= " & FunNuloData(DtpFimDeposito.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODOPAGAMENTO = "Depósito: " & Format$(DtpInicioDeposito.Value, "dd/mm/yy") & " a " & Format$(DtpFimDeposito.Value, "dd/mm/yy")
    ElseIf (Not IsNull(DtpInicioDeposito.Value)) And (IsNull(DtpFimDeposito.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelExtratoContrato.titu_dt_Deposito >= " & FunNuloData(DtpInicioDeposito.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODOPAGAMENTO = "Depósito: A partir de " & Format$(DtpInicioDeposito.Value, "dd/mm/yy")
    ElseIf (IsNull(DtpInicioDeposito.Value)) And (Not IsNull(DtpFimDeposito.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelExtratoContrato.titu_dt_Deposito <= " & FunNuloData(DtpFimDeposito.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODOPAGAMENTO = "Depósito: Até " & Format$(DtpFimDeposito.Value, "dd/mm/yy")
    Else
        XLT_FILTROPERIODOPAGAMENTO = "Depósito: Geral"
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
            " OR ConsCAPRelExtratoContrato.cont_tx_Status = '" & XLT_STATUS2 & "')"
            XGT_CONJUNCAO = " AND "
        Else
            XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "ConsCAPRelExtratoContrato.cont_tx_Status = '" & XLT_STATUS & "'"
            XGT_CONJUNCAO = " AND "
        End If
    End If
    
    
    'Desconto
    If ChkDesconto.Value = 1 Then
      XLT_FILTRODESCONTO = "1"
    Else
      XLT_FILTRODESCONTO = "0"
    End If
    
    'Seguro
    If ChkSeguro.Value = 1 Then
      XLT_FILTROSEGURO = "1"
    Else
      XLT_FILTROSEGURO = "0"
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
    
    'Filtra Títulos Pagos
    XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + "ConsCAPRelExtratoContrato.titu_dt_Pagamento is not null"
    
    'Filtra a empresa Ativa
    XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + "ConsCAPRelExtratoContrato.empr_cd_Empresa = " & CStr(PCodEmpresa)
  
    'Ordenação
    XGT_SELECAO = XGT_SELECAO + " ORDER BY titu_dt_Pagamento"
    
    'Abre os recordsets e verifica se eles possuem dados
    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute(XLT_SQL & XGT_SELECAO)
    If XGO_RSRELATORIO.EOF Then
        MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
        FrmRelTitulosLiquidadosNatureza.MousePointer = vbDefault
        Exit Sub
    End If
    Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
       
    FrmVisRelatorios.Caption = "Relatório Títulos Liquidados Por Natureza"
         
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelTitulosLiquidadosPorNatureza.rpt")
        
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
  
    'Ativa os Relatórios
    With XGR_RELATORIO.FormulaFields
        .GetItemByName("FiltroEmpreendimento").Text = Chr(34) & XLT_FILTROCONTRATO & Chr(34)
        .GetItemByName("FiltroPeriodo").Text = Chr(34) & XLT_FILTROPERIODO & Chr(34)
        .GetItemByName("FiltroPeriodoPagamento").Text = Chr(34) & XLT_FILTROPERIODOPAGAMENTO & Chr(34)
        .GetItemByName("FiltroStatus").Text = Chr(34) & XLT_FILTROSTATUS & Chr(34)
        .GetItemByName("FiltroMulta").Text = Chr(34) & XLT_FILTROMULTA & Chr(34)
        .GetItemByName("FiltroJuros").Text = Chr(34) & XLT_FILTROJUROS & Chr(34)
        .GetItemByName("FiltroOutros").Text = Chr(34) & XLT_FILTROOUTROS & Chr(34)
        .GetItemByName("FiltroSeguro").Text = Chr(34) & XLT_FILTROSEGURO & Chr(34)
        .GetItemByName("FiltroDesconto").Text = Chr(34) & XLT_FILTRODESCONTO & Chr(34)
        .GetItemByName("FiltroDesagio").Text = Chr(34) & XLT_FILTRODESAGIO & Chr(34)
        .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
    End With
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
        
    FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
    FrmRelTitulosLiquidadosNatureza.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelTitulosLiquidadosNatureza
End Sub

Private Sub Form_Load()
  
  subManutencaoJanelasAtivas "I", "FrmRelTitulosLiquidadosNatureza"
  'subConectarControleDadosNV DatTipoPlano, "SELECT * FROM TiposPlanos ORDER BY tipl_tx_Descricao", Estatico
     
  'DtpDtBase.Value = Date
  
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
