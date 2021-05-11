VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRelSegurosAPagar 
   Caption         =   "Seguros A Pagar"
   ClientHeight    =   3420
   ClientLeft      =   945
   ClientTop       =   2190
   ClientWidth     =   4620
   Icon            =   "FrmRelSegurosAPagar.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3420
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   0
      TabIndex        =   2
      Top             =   -60
      Width           =   4575
      Begin VB.Frame FraImovel 
         Height          =   555
         Left            =   60
         TabIndex        =   12
         Top             =   120
         Width           =   4455
         Begin MSMask.MaskEdBox MskImovelInicial 
            Height          =   315
            Left            =   960
            TabIndex        =   13
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
            TabIndex        =   14
            Top             =   180
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   9
            Mask            =   "AAAA.AAAA"
            PromptChar      =   " "
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
            TabIndex        =   16
            Top             =   240
            Width           =   705
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
            TabIndex        =   15
            Top             =   240
            Width           =   195
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
         TabIndex        =   8
         Top             =   1380
         Width           =   4455
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
         Height          =   675
         Left            =   60
         TabIndex        =   6
         Top             =   2100
         Width           =   4455
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
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Value           =   1  'Checked
            Width           =   2955
         End
      End
      Begin VB.Frame FraMesAno 
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
         Left            =   60
         TabIndex        =   3
         Top             =   720
         Width           =   4455
         Begin MSComCtl2.DTPicker DtpDtBase 
            Height          =   315
            Left            =   2220
            TabIndex        =   4
            Top             =   180
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MM/yyyy"
            Format          =   49086467
            CurrentDate     =   37636
         End
         Begin VB.Label LblMesAno 
            Caption         =   "Mês de Vencimento:"
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
            TabIndex        =   5
            Top             =   240
            Width           =   1815
         End
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir..."
      Height          =   345
      Left            =   2280
      TabIndex        =   0
      Top             =   3000
      Width           =   1035
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3480
      TabIndex        =   1
      Top             =   3000
      Width           =   1035
   End
End
Attribute VB_Name = "FrmRelSegurosAPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
  
  Dim XLT_FILTROIMOVEL  As String
  Dim XLT_INDEXADOR As String
  Dim XLT_SQL As String
  
  FrmRelSegurosAPagar.MousePointer = vbHourglass
  
  XGT_SELECAO = ""
  XGT_CONJUNCAO = ""
  
  If NomeSgbd = "Access" Then
  
    XLT_SQL = "" & _
    "SELECT ConsCAPRelSegurosAPAgar.*," & _
    "IIF(Indexacao1='M'" & _
      " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
      " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & "))" & _
    " as Cotacao1," & _
    "IIF(Indexacao2='M'" & _
      " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
      " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & "))" & _
    "  as Cotacao2,"
    
'    XLT_SQL = XLT_SQL & _
'    "IIF(Indexacao1='M'" & _
'      " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")))" & _
'    " as Cotacao1_MesAnterior," & _
'    "IIF(Indexacao2='M'" & _
'      " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")))" & _
'    " as Cotacao2_MesAnterior"
    
    XLT_SQL = XLT_SQL & _
    "IIF(Indexacao1='M'" & _
      " ,IIF (cont_tx_CorrecProrata = 'A',(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
      " ,IIF (cont_tx_CorrecProrata = 'P',(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")))))" & _
    " as Cotacao1_MesAnterior," & _
    "IIF(Indexacao2='M'" & _
      " ,IIF (cont_tx_CorrecProrata = 'A',(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
      " ,IIF (cont_tx_CorrecProrata = 'P',(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")))))" & _
    " as Cotacao2_MesAnterior"
  
   ' Data de Vencimento do Titulo escolhido para pagar o seguro
   XLT_SQL = XLT_SQL & _
    ",IIF(Isnull(" & _
      " (SELECT top 1 ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento FROM ConsCAPRelSegurosAPAgar as ConsCAPRelSegurosAPAgar2" & _
      " WHERE ConsCAPRelSegurosAPAgar2.titu_tx_NaturezaPlano<>'F'" & _
      " AND month(ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
      " AND year(ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
      " AND left(ConsCAPRelSegurosAPAgar2.titulo,12)=left(ConsCAPRelSegurosAPAgar.titulo,12)" & _
      " AND ConsCAPRelSegurosAPAgar2.titu_dt_Pagamento is null" & _
      " AND ConsCAPRelSegurosAPAgar2.tipl_tx_Sigla='M'" & _
      " AND ConsCAPRelSegurosAPAgar2.empr_cd_Empresa=" & PCodEmpresa & _
      " AND ConsCAPRelSegurosAPAgar2.titu_dt_Criacao=" & _
        "(SELECT top 1 min(ConsCAPRelSegurosAPAgar3.titu_dt_Criacao)" & _
        " FROM ConsCAPRelSegurosAPAgar as ConsCAPRelSegurosAPAgar3" & _
        " WHERE ConsCAPRelSegurosAPAgar3.titu_tx_NaturezaPlano<>'F'" & _
        " AND month(ConsCAPRelSegurosAPAgar3.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
        " AND year(ConsCAPRelSegurosAPAgar3.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
        " AND ConsCAPRelSegurosAPAgar3.titu_dt_Pagamento is null" & _
        " AND ConsCAPRelSegurosAPAgar3.tipl_tx_Sigla='M'" & _
        " AND ConsCAPRelSegurosAPAgar3.empr_cd_Empresa=" & PCodEmpresa & _
        " AND left(ConsCAPRelSegurosAPAgar3.titulo,12)=left(ConsCAPRelSegurosAPAgar2.titulo,12) ))),"
    
    XLT_SQL = XLT_SQL & _
      " (SELECT top 1 ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento FROM ConsCAPRelSegurosAPAgar as ConsCAPRelSegurosAPAgar2" & _
      " WHERE ConsCAPRelSegurosAPAgar2.titu_tx_NaturezaPlano<>'F'" & _
      " AND month(ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
      " AND year(ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
      " AND left(ConsCAPRelSegurosAPAgar2.titulo,12)=left(ConsCAPRelSegurosAPAgar.titulo,12)" & _
      " AND ConsCAPRelSegurosAPAgar2.titu_dt_Pagamento is null" & _
      " AND ConsCAPRelSegurosAPAgar2.empr_cd_Empresa=" & PCodEmpresa & _
      " AND ConsCAPRelSegurosAPAgar2.titu_dt_Criacao=" & _
        "(SELECT top 1 min(ConsCAPRelSegurosAPAgar3.titu_dt_Criacao)" & _
        " FROM ConsCAPRelSegurosAPAgar as ConsCAPRelSegurosAPAgar3" & _
        " WHERE ConsCAPRelSegurosAPAgar3.titu_tx_NaturezaPlano<>'F'" & _
        " AND month(ConsCAPRelSegurosAPAgar3.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
        " AND year(ConsCAPRelSegurosAPAgar3.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
        " AND ConsCAPRelSegurosAPAgar3.titu_dt_Pagamento is null" & _
        " AND ConsCAPRelSegurosAPAgar3.empr_cd_Empresa=" & PCodEmpresa & _
        " AND left(ConsCAPRelSegurosAPAgar3.titulo,12)=left(ConsCAPRelSegurosAPAgar2.titulo,12) )),"

     XLT_SQL = XLT_SQL & _
      " (SELECT top 1 ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento FROM ConsCAPRelSegurosAPAgar as ConsCAPRelSegurosAPAgar2" & _
      " WHERE ConsCAPRelSegurosAPAgar2.titu_tx_NaturezaPlano<>'F'" & _
      " AND month(ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
      " AND year(ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
      " AND left(ConsCAPRelSegurosAPAgar2.titulo,12)=left(ConsCAPRelSegurosAPAgar.titulo,12)" & _
      " AND ConsCAPRelSegurosAPAgar2.titu_dt_Pagamento is null" & _
      " AND ConsCAPRelSegurosAPAgar2.tipl_tx_Sigla='M'" & _
      " AND ConsCAPRelSegurosAPAgar2.empr_cd_Empresa=" & PCodEmpresa & _
      " AND ConsCAPRelSegurosAPAgar2.titu_dt_Criacao=" & _
        "(SELECT top 1 min(ConsCAPRelSegurosAPAgar3.titu_dt_Criacao)" & _
        " FROM ConsCAPRelSegurosAPAgar as ConsCAPRelSegurosAPAgar3" & _
        " WHERE ConsCAPRelSegurosAPAgar3.titu_tx_NaturezaPlano<>'F'" & _
        " AND month(ConsCAPRelSegurosAPAgar3.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
        " AND year(ConsCAPRelSegurosAPAgar3.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
        " AND ConsCAPRelSegurosAPAgar3.titu_dt_Pagamento is null" & _
        " AND ConsCAPRelSegurosAPAgar3.tipl_tx_Sigla='M'" & _
        " AND ConsCAPRelSegurosAPAgar3.empr_cd_Empresa=" & PCodEmpresa & _
        " AND left(ConsCAPRelSegurosAPAgar3.titulo,12)=left(ConsCAPRelSegurosAPAgar2.titulo,12) )))" & _
    "as DataBaseDesagio,"
    
  ' Código do Titulo escolhido para pagar o seguro
    XLT_SQL = XLT_SQL & _
    "IIF(Isnull(" & _
      " (SELECT top 1 ConsCAPRelSegurosAPAgar2.Titulo FROM ConsCAPRelSegurosAPAgar as ConsCAPRelSegurosAPAgar2" & _
      " WHERE ConsCAPRelSegurosAPAgar2.titu_tx_NaturezaPlano<>'F'" & _
      " AND month(ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
      " AND year(ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
      " AND left(ConsCAPRelSegurosAPAgar2.titulo,12)=left(ConsCAPRelSegurosAPAgar.titulo,12)" & _
      " AND ConsCAPRelSegurosAPAgar2.titu_dt_Pagamento is null" & _
      " AND ConsCAPRelSegurosAPAgar2.tipl_tx_Sigla='M'" & _
      " AND ConsCAPRelSegurosAPAgar2.empr_cd_Empresa=" & PCodEmpresa & _
      " AND ConsCAPRelSegurosAPAgar2.titu_dt_Criacao=" & _
        "(SELECT top 1 min(ConsCAPRelSegurosAPAgar3.titu_dt_Criacao)" & _
        " FROM ConsCAPRelSegurosAPAgar as ConsCAPRelSegurosAPAgar3" & _
        " WHERE ConsCAPRelSegurosAPAgar3.titu_tx_NaturezaPlano<>'F'" & _
        " AND month(ConsCAPRelSegurosAPAgar3.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
        " AND year(ConsCAPRelSegurosAPAgar3.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
        " AND ConsCAPRelSegurosAPAgar3.titu_dt_Pagamento is null" & _
        " AND ConsCAPRelSegurosAPAgar3.tipl_tx_Sigla='M'" & _
        " AND ConsCAPRelSegurosAPAgar3.empr_cd_Empresa=" & PCodEmpresa & _
        " AND left(ConsCAPRelSegurosAPAgar3.titulo,12)=left(ConsCAPRelSegurosAPAgar2.titulo,12) ))),"

    XLT_SQL = XLT_SQL & _
      " (SELECT top 1 ConsCAPRelSegurosAPAgar2.Titulo FROM ConsCAPRelSegurosAPAgar as ConsCAPRelSegurosAPAgar2" & _
      " WHERE ConsCAPRelSegurosAPAgar2.titu_tx_NaturezaPlano<>'F'" & _
      " AND month(ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
      " AND year(ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
      " AND left(ConsCAPRelSegurosAPAgar2.titulo,12)=left(ConsCAPRelSegurosAPAgar.titulo,12)" & _
      " AND ConsCAPRelSegurosAPAgar2.titu_dt_Pagamento is null" & _
      " AND ConsCAPRelSegurosAPAgar2.empr_cd_Empresa=" & PCodEmpresa & _
      " AND ConsCAPRelSegurosAPAgar2.titu_dt_Criacao=" & _
        "(SELECT top 1 min(ConsCAPRelSegurosAPAgar3.titu_dt_Criacao)" & _
        " FROM ConsCAPRelSegurosAPAgar as ConsCAPRelSegurosAPAgar3" & _
        " WHERE ConsCAPRelSegurosAPAgar3.titu_tx_NaturezaPlano<>'F'" & _
        " AND month(ConsCAPRelSegurosAPAgar3.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
        " AND year(ConsCAPRelSegurosAPAgar3.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
        " AND ConsCAPRelSegurosAPAgar3.titu_dt_Pagamento is null" & _
        " AND ConsCAPRelSegurosAPAgar3.empr_cd_Empresa=" & PCodEmpresa & _
        " AND left(ConsCAPRelSegurosAPAgar3.titulo,12)=left(ConsCAPRelSegurosAPAgar2.titulo,12) )),"

     XLT_SQL = XLT_SQL & _
      " (SELECT top 1 ConsCAPRelSegurosAPAgar2.Titulo FROM ConsCAPRelSegurosAPAgar as ConsCAPRelSegurosAPAgar2" & _
      " WHERE ConsCAPRelSegurosAPAgar2.titu_tx_NaturezaPlano<>'F'" & _
      " AND month(ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
      " AND year(ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
      " AND left(ConsCAPRelSegurosAPAgar2.titulo,12)=left(ConsCAPRelSegurosAPAgar.titulo,12)" & _
      " AND ConsCAPRelSegurosAPAgar2.titu_dt_Pagamento is null" & _
      " AND ConsCAPRelSegurosAPAgar2.tipl_tx_Sigla='M'" & _
      " AND ConsCAPRelSegurosAPAgar2.empr_cd_Empresa=" & PCodEmpresa & _
      " AND ConsCAPRelSegurosAPAgar2.titu_dt_Criacao=" & _
        "(SELECT top 1 min(ConsCAPRelSegurosAPAgar3.titu_dt_Criacao)" & _
        " FROM ConsCAPRelSegurosAPAgar as ConsCAPRelSegurosAPAgar3" & _
        " WHERE ConsCAPRelSegurosAPAgar3.titu_tx_NaturezaPlano<>'F'" & _
        " AND month(ConsCAPRelSegurosAPAgar3.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
        " AND year(ConsCAPRelSegurosAPAgar3.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
        " AND ConsCAPRelSegurosAPAgar3.titu_dt_Pagamento is null" & _
        " AND ConsCAPRelSegurosAPAgar3.tipl_tx_Sigla='M'" & _
        " AND ConsCAPRelSegurosAPAgar3.empr_cd_Empresa=" & PCodEmpresa & _
        " AND left(ConsCAPRelSegurosAPAgar3.titulo,12)=left(ConsCAPRelSegurosAPAgar2.titulo,12) )))" & _
    "as TituloSeguro" & _
    " From ConsCAPRelSegurosAPAgar WHERE "
    
    XLT_SQL = XLT_SQL & _
    "(SELECT Titulos.titu_dt_Pagamento FROM Titulos WHERE " & _
      " month(Titulos.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
      " AND year(Titulos.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
      " AND Titulos.empr_cd_Empresa=" & PCodEmpresa & _
      " AND left(ConsCAPRelSegurosAPAgar.titulo,12)= Titulos.empd_cd_Empreendimento + '.'" & _
      " + Titulos.imov_cd_Imovel + '.' + Titulos.cont_cd_Contrato" & _
      " AND Titulos.titu_dt_Pagamento is not null" & _
      " AND Titulos.titu_vl_Seguro>0) is null "
  
  Else
  
    XLT_SQL = "" & _
    "SELECT ConsCAPRelSegurosAPAgar.*," & _
    "CASE Indexacao1   when 'M' then" & _
      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
      " when 'D' then" & _
      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
    " end as Cotacao1," & _
    "CASE  Indexacao2 when 'M' then" & _
      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
      " when 'D'  then" & _
      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
    " end as Cotacao2,"
  
'    XLT_SQL = XLT_SQL & _
'    "CASE  Indexacao1 when 'M' then" & _
'      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
'      " when 'D' then Null " & _
'    " end as Cotacao1_MesAnterior," & _
'    "CASE Indexacao2 when 'M' then" & _
'      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
'      " when 'D' then Null " & _
'    " end As Cotacao2_MesAnterior"
    
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
    " end As Cotacao2_MesAnterior"
    
   ' Data de Vencimento do Titulo escolhido para pagar o seguro
    XLT_SQL = XLT_SQL & _
    ",Isnull(" & _
      " (SELECT top 1 ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento FROM ConsCAPRelSegurosAPAgar as ConsCAPRelSegurosAPAgar2" & _
      " WHERE ConsCAPRelSegurosAPAgar2.titu_tx_NaturezaPlano<>'F'" & _
      " AND month(ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
      " AND year(ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
      " AND left(ConsCAPRelSegurosAPAgar2.titulo,12)=left(ConsCAPRelSegurosAPAgar.titulo,12)" & _
      " AND ConsCAPRelSegurosAPAgar2.titu_dt_Pagamento is null" & _
      " AND ConsCAPRelSegurosAPAgar2.tipl_tx_Sigla='M'" & _
      " AND ConsCAPRelSegurosAPAgar2.empr_cd_Empresa=" & PCodEmpresa & _
      " AND ConsCAPRelSegurosAPAgar2.titu_dt_Criacao=" & _
        "(SELECT top 1 min(ConsCAPRelSegurosAPAgar3.titu_dt_Criacao)" & _
        " FROM ConsCAPRelSegurosAPAgar as ConsCAPRelSegurosAPAgar3" & _
        " WHERE ConsCAPRelSegurosAPAgar3.titu_tx_NaturezaPlano<>'F'" & _
        " AND month(ConsCAPRelSegurosAPAgar3.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
        " AND year(ConsCAPRelSegurosAPAgar3.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
        " AND ConsCAPRelSegurosAPAgar3.titu_dt_Pagamento is null" & _
        " AND ConsCAPRelSegurosAPAgar3.tipl_tx_Sigla='M'" & _
        " AND ConsCAPRelSegurosAPAgar3.empr_cd_Empresa=" & PCodEmpresa & _
        " AND left(ConsCAPRelSegurosAPAgar3.titulo,12)=left(ConsCAPRelSegurosAPAgar2.titulo,12) )),"
    
    XLT_SQL = XLT_SQL & _
      " (SELECT top 1 ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento FROM ConsCAPRelSegurosAPAgar as ConsCAPRelSegurosAPAgar2" & _
      " WHERE ConsCAPRelSegurosAPAgar2.titu_tx_NaturezaPlano<>'F'" & _
      " AND month(ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
      " AND year(ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
      " AND left(ConsCAPRelSegurosAPAgar2.titulo,12)=left(ConsCAPRelSegurosAPAgar.titulo,12)" & _
      " AND ConsCAPRelSegurosAPAgar2.titu_dt_Pagamento is null" & _
      " AND ConsCAPRelSegurosAPAgar2.empr_cd_Empresa=" & PCodEmpresa & _
      " AND ConsCAPRelSegurosAPAgar2.titu_dt_Criacao=" & _
        "(SELECT top 1 min(ConsCAPRelSegurosAPAgar3.titu_dt_Criacao)" & _
        " FROM ConsCAPRelSegurosAPAgar as ConsCAPRelSegurosAPAgar3" & _
        " WHERE ConsCAPRelSegurosAPAgar3.titu_tx_NaturezaPlano<>'F'" & _
        " AND month(ConsCAPRelSegurosAPAgar3.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
        " AND year(ConsCAPRelSegurosAPAgar3.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
        " AND ConsCAPRelSegurosAPAgar3.titu_dt_Pagamento is null" & _
        " AND ConsCAPRelSegurosAPAgar3.empr_cd_Empresa=" & PCodEmpresa & _
        " AND left(ConsCAPRelSegurosAPAgar3.titulo,12)=left(ConsCAPRelSegurosAPAgar2.titulo,12) )))" & _
    "as DataBaseDesagio"
    
   ' Código do Titulo escolhido para pagar o seguro
    XLT_SQL = XLT_SQL & _
    ",Isnull(" & _
      " (SELECT top 1 ConsCAPRelSegurosAPAgar2.Titulo FROM ConsCAPRelSegurosAPAgar as ConsCAPRelSegurosAPAgar2" & _
      " WHERE ConsCAPRelSegurosAPAgar2.titu_tx_NaturezaPlano<>'F'" & _
      " AND month(ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
      " AND year(ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
      " AND left(ConsCAPRelSegurosAPAgar2.titulo,12)=left(ConsCAPRelSegurosAPAgar.titulo,12)" & _
      " AND ConsCAPRelSegurosAPAgar2.titu_dt_Pagamento is null" & _
      " AND ConsCAPRelSegurosAPAgar2.tipl_tx_Sigla='M'" & _
      " AND ConsCAPRelSegurosAPAgar2.empr_cd_Empresa=" & PCodEmpresa & _
      " AND ConsCAPRelSegurosAPAgar2.titu_dt_Criacao=" & _
        "(SELECT top 1 min(ConsCAPRelSegurosAPAgar3.titu_dt_Criacao)" & _
        " FROM ConsCAPRelSegurosAPAgar as ConsCAPRelSegurosAPAgar3" & _
        " WHERE ConsCAPRelSegurosAPAgar3.titu_tx_NaturezaPlano<>'F'" & _
        " AND month(ConsCAPRelSegurosAPAgar3.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
        " AND year(ConsCAPRelSegurosAPAgar3.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
        " AND ConsCAPRelSegurosAPAgar3.titu_dt_Pagamento is null" & _
        " AND ConsCAPRelSegurosAPAgar3.tipl_tx_Sigla='M'" & _
        " AND ConsCAPRelSegurosAPAgar3.empr_cd_Empresa=" & PCodEmpresa & _
        " AND left(ConsCAPRelSegurosAPAgar3.titulo,12)=left(ConsCAPRelSegurosAPAgar2.titulo,12) )),"
    
    XLT_SQL = XLT_SQL & _
      " (SELECT top 1 ConsCAPRelSegurosAPAgar2.Titulo FROM ConsCAPRelSegurosAPAgar as ConsCAPRelSegurosAPAgar2" & _
      " WHERE ConsCAPRelSegurosAPAgar2.titu_tx_NaturezaPlano<>'F'" & _
      " AND month(ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
      " AND year(ConsCAPRelSegurosAPAgar2.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
      " AND left(ConsCAPRelSegurosAPAgar2.titulo,12)=left(ConsCAPRelSegurosAPAgar.titulo,12)" & _
      " AND ConsCAPRelSegurosAPAgar2.titu_dt_Pagamento is null" & _
      " AND ConsCAPRelSegurosAPAgar2.empr_cd_Empresa=" & PCodEmpresa & _
      " AND ConsCAPRelSegurosAPAgar2.titu_dt_Criacao=" & _
        "(SELECT top 1 min(ConsCAPRelSegurosAPAgar3.titu_dt_Criacao)" & _
        " FROM ConsCAPRelSegurosAPAgar as ConsCAPRelSegurosAPAgar3" & _
        " WHERE ConsCAPRelSegurosAPAgar3.titu_tx_NaturezaPlano<>'F'" & _
        " AND month(ConsCAPRelSegurosAPAgar3.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
        " AND year(ConsCAPRelSegurosAPAgar3.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
        " AND ConsCAPRelSegurosAPAgar3.titu_dt_Pagamento is null" & _
        " AND ConsCAPRelSegurosAPAgar3.empr_cd_Empresa=" & PCodEmpresa & _
        " AND left(ConsCAPRelSegurosAPAgar3.titulo,12)=left(ConsCAPRelSegurosAPAgar2.titulo,12) )))" & _
    "as TituloSeguro" & _
    " From ConsCAPRelSegurosAPAgar WHERE "
    
    XLT_SQL = XLT_SQL & _
    "(SELECT Titulos.titu_dt_Pagamento FROM Titulos WHERE " & _
      " month(Titulos.titu_dt_Vencimento) = " & Month(DtpDtBase.Value) & _
      " AND year(Titulos.titu_dt_Vencimento) = " & Year(DtpDtBase.Value) & _
      " AND Titulos.empr_cd_Empresa=" & PCodEmpresa & _
      " AND left(ConsCAPRelSegurosAPAgar.titulo,12)= Titulos.empd_cd_Empreendimento + '.'" & _
      " + Titulos.imov_cd_Imovel + '.' + Titulos.cont_cd_Contrato" & _
      " AND Titulos.titu_dt_Pagamento is not null" & _
      " AND Titulos.titu_vl_Seguro>0) is null "
  
  End If
  
  XGT_CONJUNCAO = " AND "
  
  'Imóvel
  If MskImovelInicial.Text <> "    .    " And MskImovelFinal.Text <> "    .    " Then
    XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
    "ConsCAPRelSegurosAPagar.Imovel >= '" & MskImovelInicial.Text & "'" & _
    " AND ConsCAPRelSegurosAPagar.Imovel <= '" & MskImovelFinal.Text & "'"
    XGT_CONJUNCAO = " AND "
    XLT_FILTROIMOVEL = "Imóvel: " & MskImovelInicial & " a " & MskImovelFinal
  ElseIf MskImovelInicial.Text <> "    .    " And MskImovelFinal.Text = "    .    " Then
    XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
    "ConsCAPRelSegurosAPagar.Imovel >= '" & MskImovelInicial.Text & "'"
    XGT_CONJUNCAO = " AND "
    XLT_FILTROIMOVEL = "Imóvel: A partir de " & MskImovelInicial
  ElseIf MskImovelInicial.Text = "    .    " And MskImovelFinal.Text <> "    .    " Then
    XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
    "ConsCAPRelSegurosAPagar.Imovel <= '" & MskImovelFinal.Text & "'"
    XGT_CONJUNCAO = " AND "
    XLT_FILTROIMOVEL = "Imóvel: Até " & MskImovelInicial
  Else
    XLT_FILTROIMOVEL = "Imóvel: Geral"
  End If
  
  'Indexador
  If OptValorHistorico.Value = True Then
    XLT_INDEXADOR = 0
  ElseIf OptIndexador1.Value = True Then
    XLT_INDEXADOR = 1
  ElseIf OptIndexador2.Value = True Then
    XLT_INDEXADOR = 2
  End If
  
  'Filtra Títulos não Pagos
  XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + "ConsCAPRelSegurosAPAgar.titu_dt_Pagamento is null"
  XGT_CONJUNCAO = " AND "
  
  'Filtra a empresa Ativa
  XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + "ConsCAPRelSegurosAPAgar.empr_cd_Empresa = " & CStr(PCodEmpresa)

  'Ordenação
  'XGT_SELECAO = XGT_SELECAO + " ORDER BY titu_dt_Pagamento"
  
  'Abre os recordsets e verifica se eles possuem dados
  Set XGO_RSRELATORIO = ConexaoRelatorio.Execute(XLT_SQL & XGT_SELECAO)
  XGO_RSRELATORIO.Filter = "DataBaseDesagio <> null"
  If XGO_RSRELATORIO.EOF Then
      MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
      FrmRelSegurosAPagar.MousePointer = vbDefault
      Exit Sub
  End If
  
  Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
      
  If ChkEmpreendimentos.Value = 1 Then
    'Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelSegurosAPagarAgrupados.rpt")
    Set XGR_RELATORIO = RptSegurosAPagarAgrupados
    FrmVisRelatorios.Caption = "Relatório Seguros A Pagar - Empreendimentos Agrupados"
  Else
    'Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelSegurosAPagar.rpt")
    Set XGR_RELATORIO = RptSegurosAPagar
    FrmVisRelatorios.Caption = "Relatório Seguros A Pagar"
  End If
  XGR_RELATORIO.DiscardSavedData
  
  'Passa o recordset para o relatório
  XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO

  'Ativa os Relatórios
  With XGR_RELATORIO.FormulaFields
      .GetItemByName("FiltroImovel").Text = Chr(34) & XLT_FILTROIMOVEL & Chr(34)
      .GetItemByName("FiltroDataBase").Text = Chr(34) & Format(DtpDtBase.Value, "\01/mm/yy") & Chr(34)
      .GetItemByName("Indexador").Text = Chr(34) & XLT_INDEXADOR & Chr(34)
      .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
  End With
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
  
  FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
  FrmRelSegurosAPagar.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelSegurosAPagar
End Sub

Private Sub Form_Load()
  
  subManutencaoJanelasAtivas "I", "FrmRelSegurosAPAgar"
  
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
