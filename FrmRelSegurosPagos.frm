VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRelSegurosPagos 
   Caption         =   "Seguros Pagos"
   ClientHeight    =   2790
   ClientLeft      =   945
   ClientTop       =   2190
   ClientWidth     =   4620
   Icon            =   "FrmRelSegurosPagos.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2790
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2235
      Left            =   0
      TabIndex        =   2
      Top             =   -60
      Width           =   4575
      Begin VB.Frame FraImovel 
         Height          =   555
         Left            =   60
         TabIndex        =   10
         Top             =   120
         Width           =   4455
         Begin MSMask.MaskEdBox MskImovelInicial 
            Height          =   315
            Left            =   960
            TabIndex        =   11
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
            TabIndex        =   12
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
            TabIndex        =   14
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
            TabIndex        =   13
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
         TabIndex        =   6
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
            TabIndex        =   9
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
            TabIndex        =   8
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
            TabIndex        =   7
            Top             =   300
            Width           =   1575
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
      Top             =   2340
      Width           =   1035
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3480
      TabIndex        =   1
      Top             =   2340
      Width           =   1035
   End
End
Attribute VB_Name = "FrmRelSegurosPagos"
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
  
  FrmRelSegurosPagos.MousePointer = vbHourglass
  
  XGT_SELECAO = ""
  XGT_CONJUNCAO = ""
  
  If NomeSgbd = "Access" Then
    
    XLT_SQL = "" & _
    "SELECT ConsCAPRelSegurosPagos.*," & _
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
'    " as Cotacao2_MesAnterior,"

    XLT_SQL = XLT_SQL & _
    "IIF(Indexacao1='M'" & _
      " ,IIF (cont_tx_CorrecProrata = 'A',(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
      " ,IIF (cont_tx_CorrecProrata = 'P',(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")))))" & _
    " as Cotacao1_MesAnterior," & _
    "IIF(Indexacao2='M'" & _
      " ,IIF (cont_tx_CorrecProrata = 'A',(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
      " ,IIF (cont_tx_CorrecProrata = 'P',(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")))))" & _
    " as Cotacao2_MesAnterior,"
     
 ' Data de Vencimento Do Seguro Pago
     XLT_SQL = XLT_SQL & _
    " IIF(Isnull((SELECT  Titulos.titu_dt_Vencimento FROM Titulos" & _
    " WHERE Titulos.titu_tx_NaturezaPlano<>'S' AND month(Titulos.titu_dt_Pagamento) = " & Month(DtpDtBase.Value) & _
    " AND year(Titulos.titu_dt_Pagamento) = " & Year(DtpDtBase.Value) & _
    " AND left(ConsCAPRelSegurosPagos.titulo,12)=Titulos.empd_cd_Empreendimento + '.' + Titulos.imov_cd_Imovel + '.' + Titulos.cont_cd_Contrato" & _
    " AND Titulos.titu_dt_Pagamento is not null" & _
    " AND Titulos.empr_cd_Empresa=" & PCodEmpresa & _
    " AND Titulos.titu_vl_Seguro>0)),"
    
    XLT_SQL = XLT_SQL & _
    " (SELECT  Titulos.titu_dt_Vencimento FROM Titulos" & _
    " WHERE Titulos.titu_tx_NaturezaPlano='S' AND month(Titulos.titu_dt_Pagamento) = " & Month(DtpDtBase.Value) & _
    " AND year(Titulos.titu_dt_Pagamento) = " & Year(DtpDtBase.Value) & _
    " AND left(ConsCAPRelSegurosPagos.titulo,12)=Titulos.empd_cd_Empreendimento + '.' + Titulos.imov_cd_Imovel + '.' + Titulos.cont_cd_Contrato" & _
    " AND Titulos.titu_dt_Pagamento is not null" & _
    " AND Titulos.empr_cd_Empresa=" & PCodEmpresa & _
    " AND Titulos.titu_vl_Parcela>0),"
    
    XLT_SQL = XLT_SQL & _
    " (SELECT  Titulos.titu_dt_Vencimento FROM Titulos" & _
    " WHERE Titulos.titu_tx_NaturezaPlano<>'S' AND month(Titulos.titu_dt_Pagamento) = " & Month(DtpDtBase.Value) & _
    " AND year(Titulos.titu_dt_Pagamento) = " & Year(DtpDtBase.Value) & _
    " AND left(ConsCAPRelSegurosPagos.titulo,12)=Titulos.empd_cd_Empreendimento + '.' + Titulos.imov_cd_Imovel + '.' + Titulos.cont_cd_Contrato" & _
    " AND Titulos.titu_dt_Pagamento is not null" & _
    " AND Titulos.empr_cd_Empresa=" & PCodEmpresa & _
    " AND Titulos.titu_vl_Seguro>0))" & _
    " as DataBaseDesagio,"
    
' Data de Pagamento do Seguro Pago
    
    XLT_SQL = XLT_SQL & _
    " IIF(Isnull((SELECT  Titulos.titu_dt_Pagamento FROM Titulos" & _
    " WHERE Titulos.titu_tx_NaturezaPlano<>'S' AND month(Titulos.titu_dt_Pagamento) = " & Month(DtpDtBase.Value) & _
    " AND year(Titulos.titu_dt_Pagamento) = " & Year(DtpDtBase.Value) & _
    " AND left(ConsCAPRelSegurosPagos.titulo,12)=Titulos.empd_cd_Empreendimento + '.' + Titulos.imov_cd_Imovel + '.' + Titulos.cont_cd_Contrato" & _
    " AND Titulos.titu_dt_Pagamento is not null" & _
    " AND Titulos.empr_cd_Empresa=" & PCodEmpresa & _
    " AND Titulos.titu_vl_Seguro>0)),"
    
    XLT_SQL = XLT_SQL & _
    " (SELECT  Titulos.titu_dt_Pagamento FROM Titulos" & _
    " WHERE Titulos.titu_tx_NaturezaPlano='S' AND month(Titulos.titu_dt_Pagamento) = " & Month(DtpDtBase.Value) & _
    " AND year(Titulos.titu_dt_Pagamento) = " & Year(DtpDtBase.Value) & _
    " AND left(ConsCAPRelSegurosPagos.titulo,12)=Titulos.empd_cd_Empreendimento + '.' + Titulos.imov_cd_Imovel + '.' + Titulos.cont_cd_Contrato" & _
    " AND Titulos.titu_dt_Pagamento is not null" & _
    " AND Titulos.empr_cd_Empresa=" & PCodEmpresa & _
    " AND Titulos.titu_vl_Parcela>0),"
    
    XLT_SQL = XLT_SQL & _
    " (SELECT  Titulos.titu_dt_Pagamento FROM Titulos" & _
    " WHERE Titulos.titu_tx_NaturezaPlano<>'S' AND month(Titulos.titu_dt_Pagamento) = " & Month(DtpDtBase.Value) & _
    " AND year(Titulos.titu_dt_Pagamento) = " & Year(DtpDtBase.Value) & _
    " AND left(ConsCAPRelSegurosPagos.titulo,12)=Titulos.empd_cd_Empreendimento + '.' + Titulos.imov_cd_Imovel + '.' + Titulos.cont_cd_Contrato" & _
    " AND Titulos.titu_dt_Pagamento is not null" & _
    " AND Titulos.empr_cd_Empresa=" & PCodEmpresa & _
    " AND Titulos.titu_vl_Seguro>0))" & _
    " as DataPagamento,"
    
' Valor do Seguro Pago

    XLT_SQL = XLT_SQL & _
    " IIF(Isnull((SELECT  Titulos.titu_vl_Seguro FROM Titulos" & _
    " WHERE Titulos.titu_tx_NaturezaPlano<>'S' AND month(Titulos.titu_dt_Pagamento) = " & Month(DtpDtBase.Value) & _
    " AND year(Titulos.titu_dt_Pagamento) = " & Year(DtpDtBase.Value) & _
    " AND left(ConsCAPRelSegurosPagos.titulo,12)=Titulos.empd_cd_Empreendimento + '.' + Titulos.imov_cd_Imovel + '.' + Titulos.cont_cd_Contrato" & _
    " AND Titulos.titu_dt_Pagamento is not null" & _
    " AND Titulos.empr_cd_Empresa=" & PCodEmpresa & _
    " AND Titulos.titu_vl_Seguro>0)),"
    
    XLT_SQL = XLT_SQL & _
    " (SELECT  Titulos.titu_vl_Parcela FROM Titulos" & _
    " WHERE Titulos.titu_tx_NaturezaPlano='S' AND month(Titulos.titu_dt_Pagamento) = " & Month(DtpDtBase.Value) & _
    " AND year(Titulos.titu_dt_Pagamento) = " & Year(DtpDtBase.Value) & _
    " AND left(ConsCAPRelSegurosPagos.titulo,12)=Titulos.empd_cd_Empreendimento + '.' + Titulos.imov_cd_Imovel + '.' + Titulos.cont_cd_Contrato" & _
    " AND Titulos.titu_dt_Pagamento is not null" & _
    " AND Titulos.empr_cd_Empresa=" & PCodEmpresa & _
    " AND Titulos.titu_vl_Parcela>0),"
    
    XLT_SQL = XLT_SQL & _
    " (SELECT  Titulos.titu_vl_Seguro FROM Titulos" & _
    " WHERE Titulos.titu_tx_NaturezaPlano<>'S' AND month(Titulos.titu_dt_Pagamento) = " & Month(DtpDtBase.Value) & _
    " AND year(Titulos.titu_dt_Pagamento) = " & Year(DtpDtBase.Value) & _
    " AND left(ConsCAPRelSegurosPagos.titulo,12)=Titulos.empd_cd_Empreendimento + '.' + Titulos.imov_cd_Imovel + '.' + Titulos.cont_cd_Contrato" & _
    " AND Titulos.titu_dt_Pagamento is not null" & _
    " AND Titulos.empr_cd_Empresa=" & PCodEmpresa & _
    " AND Titulos.titu_vl_Seguro>0))" & _
    " as SeguroPago"
    
''''''''''''''''''''''''''''''''''''''''''''''''
    
    XLT_SQL = XLT_SQL & _
    " From ConsCAPRelSegurosPagos WHERE " & _
    " (SELECT titulos.titu_dt_Pagamento FROM Titulos WHERE  month(titulos.titu_dt_Pagamento) = " & Month(DtpDtBase.Value) & _
    " AND year(titulos.titu_dt_Pagamento) = " & Year(DtpDtBase.Value) & _
    " AND titulos.empr_cd_Empresa=" & PCodEmpresa & _
    " AND left(ConsCAPRelSegurosPagos.titulo,12)= Titulos.empd_cd_Empreendimento + '.' + Titulos.imov_cd_Imovel + '.' + Titulos.cont_cd_Contrato" & _
    " AND titu_dt_Pagamento is not null" & _
    " AND titulos.titu_vl_Seguro>0) is not null"
  
  Else
  
    XLT_SQL = "" & _
    "SELECT ConsCAPRelSegurosPagos.*," & _
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
    
 ' Data de Vencimento do Seguro Pago
    XLT_SQL = XLT_SQL & _
    " ,Isnull((SELECT  Titulos.titu_dt_Vencimento FROM Titulos" & _
    " WHERE Titulos.titu_tx_NaturezaPlano<>'S' AND month(Titulos.titu_dt_Pagamento) = " & Month(DtpDtBase.Value) & _
    " AND year(Titulos.titu_dt_Pagamento) = " & Year(DtpDtBase.Value) & _
    " AND left(ConsCAPRelSegurosPagos.titulo,12)=Titulos.empd_cd_Empreendimento + '.' + Titulos.imov_cd_Imovel + '.' + Titulos.cont_cd_Contrato" & _
    " AND Titulos.titu_dt_Pagamento is not null" & _
    " AND Titulos.empr_cd_Empresa=" & PCodEmpresa & _
    " AND Titulos.titu_vl_Seguro>0),"
    
    XLT_SQL = XLT_SQL & _
    " (SELECT  Titulos.titu_dt_Vencimento FROM Titulos" & _
    " WHERE Titulos.titu_tx_NaturezaPlano='S' AND month(Titulos.titu_dt_Pagamento) = " & Month(DtpDtBase.Value) & _
    " AND year(Titulos.titu_dt_Pagamento) = " & Year(DtpDtBase.Value) & _
    " AND left(ConsCAPRelSegurosPagos.titulo,12)=Titulos.empd_cd_Empreendimento + '.' + Titulos.imov_cd_Imovel + '.' + Titulos.cont_cd_Contrato" & _
    " AND Titulos.titu_dt_Pagamento is not null" & _
    " AND Titulos.empr_cd_Empresa=" & PCodEmpresa & _
    " AND Titulos.titu_vl_Parcela>0))" & _
    " as DataBaseDesagio,"
    
  ' Data de PAgamento do Seguro Pago
    XLT_SQL = XLT_SQL & _
    " Isnull((SELECT  Titulos.titu_dt_Pagamento FROM Titulos" & _
    " WHERE Titulos.titu_tx_NaturezaPlano<>'S' AND month(Titulos.titu_dt_Pagamento) = " & Month(DtpDtBase.Value) & _
    " AND year(Titulos.titu_dt_Pagamento) = " & Year(DtpDtBase.Value) & _
    " AND left(ConsCAPRelSegurosPagos.titulo,12)=Titulos.empd_cd_Empreendimento + '.' + Titulos.imov_cd_Imovel + '.' + Titulos.cont_cd_Contrato" & _
    " AND Titulos.titu_dt_Pagamento is not null" & _
    " AND Titulos.empr_cd_Empresa=" & PCodEmpresa & _
    " AND Titulos.titu_vl_Seguro>0),"
    
    XLT_SQL = XLT_SQL & _
    " (SELECT  Titulos.titu_dt_Pagamento FROM Titulos" & _
    " WHERE Titulos.titu_tx_NaturezaPlano='S' AND month(Titulos.titu_dt_Pagamento) = " & Month(DtpDtBase.Value) & _
    " AND year(Titulos.titu_dt_Pagamento) = " & Year(DtpDtBase.Value) & _
    " AND left(ConsCAPRelSegurosPagos.titulo,12)=Titulos.empd_cd_Empreendimento + '.' + Titulos.imov_cd_Imovel + '.' + Titulos.cont_cd_Contrato" & _
    " AND Titulos.titu_dt_Pagamento is not null" & _
    " AND Titulos.empr_cd_Empresa=" & PCodEmpresa & _
    " AND Titulos.titu_vl_Parcela>0))" & _
    " as DataPagamento,"
    
  ' Vlor do Seguro Pago
    XLT_SQL = XLT_SQL & _
    " isnull((SELECT  Titulos.titu_vl_Seguro FROM Titulos" & _
    " WHERE Titulos.titu_tx_NaturezaPlano<>'S' AND month(Titulos.titu_dt_Pagamento) = " & Month(DtpDtBase.Value) & _
    " AND year(Titulos.titu_dt_Pagamento) = " & Year(DtpDtBase.Value) & _
    " AND left(ConsCAPRelSegurosPagos.titulo,12)=Titulos.empd_cd_Empreendimento + '.' + Titulos.imov_cd_Imovel + '.' + Titulos.cont_cd_Contrato" & _
    " AND Titulos.titu_dt_Pagamento is not null" & _
    " AND Titulos.empr_cd_Empresa=" & PCodEmpresa & _
    " AND Titulos.titu_vl_Seguro>0),"
    
    XLT_SQL = XLT_SQL & _
    " (SELECT  Titulos.titu_vl_Parcela FROM Titulos" & _
    " WHERE Titulos.titu_tx_NaturezaPlano='S' AND month(Titulos.titu_dt_Pagamento) = " & Month(DtpDtBase.Value) & _
    " AND year(Titulos.titu_dt_Pagamento) = " & Year(DtpDtBase.Value) & _
    " AND left(ConsCAPRelSegurosPagos.titulo,12)=Titulos.empd_cd_Empreendimento + '.' + Titulos.imov_cd_Imovel + '.' + Titulos.cont_cd_Contrato" & _
    " AND Titulos.titu_dt_Pagamento is not null" & _
    " AND Titulos.empr_cd_Empresa=" & PCodEmpresa & _
    " AND Titulos.titu_vl_Parcela>0))" & _
    " as SeguroPago"
    
    XLT_SQL = XLT_SQL & _
    " From ConsCAPRelSegurosPagos WHERE " & _
    " (SELECT titulos.titu_dt_Pagamento FROM Titulos WHERE  month(titulos.titu_dt_Pagamento) = " & Month(DtpDtBase.Value) & _
    " AND year(titulos.titu_dt_Pagamento) = " & Year(DtpDtBase.Value) & _
    " AND titulos.empr_cd_Empresa=" & PCodEmpresa & _
    " AND left(ConsCAPRelSegurosPagos.titulo,12)= Titulos.empd_cd_Empreendimento + '.' + Titulos.imov_cd_Imovel + '.' + Titulos.cont_cd_Contrato" & _
    " AND titu_dt_Pagamento is not null" & _
    " AND titulos.titu_vl_Seguro>0) is not null"
    
  End If
  
  XGT_CONJUNCAO = " AND "
  
  'Imóvel
  If MskImovelInicial.Text <> "    .    " And MskImovelFinal.Text <> "    .    " Then
    XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
    "ConsCAPRelSegurosPagos.Imovel >= '" & MskImovelInicial.Text & "'" & _
    " AND ConsCAPRelSegurosPagos.Imovel <= '" & MskImovelFinal.Text & "'"
    XGT_CONJUNCAO = " AND "
    XLT_FILTROIMOVEL = "Imóvel: " & MskImovelInicial & " a " & MskImovelFinal
  ElseIf MskImovelInicial.Text <> "    .    " And MskImovelFinal.Text = "    .    " Then
    XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
    "ConsCAPRelSegurosPagos.Imovel >= '" & MskImovelInicial.Text & "'"
    XGT_CONJUNCAO = " AND "
    XLT_FILTROIMOVEL = "Imóvel: A partir de " & MskImovelInicial
  ElseIf MskImovelInicial.Text = "    .    " And MskImovelFinal.Text <> "    .    " Then
    XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
    "ConsCAPRelSegurosPagos.Imovel <= '" & MskImovelFinal.Text & "'"
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
  
  'Filtra Títulos Não Pagos
  XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + "ConsCAPRelSegurosPagos.titu_dt_Pagamento is null"
  XGT_CONJUNCAO = " AND "
  
  'Filtra a empresa Ativa
  XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + "ConsCAPRelSegurosPagos.empr_cd_Empresa = " & CStr(PCodEmpresa)

  'Ordenação
  'XGT_SELECAO = XGT_SELECAO + " ORDER BY titu_dt_Pagamento"
  
  'Abre os recordsets e verifica se eles possuem dados
  Set XGO_RSRELATORIO = ConexaoRelatorio.Execute(XLT_SQL & XGT_SELECAO)
  If XGO_RSRELATORIO.EOF Then
      MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
      FrmRelSegurosPagos.MousePointer = vbDefault
      Exit Sub
  End If
  Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
  
  'Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelSegurosPagos.rpt")
  Set XGR_RELATORIO = RptSegurosPagos
  XGR_RELATORIO.DiscardSavedData
  
  FrmVisRelatorios.Caption = "Relatório Seguros Pagos"
  
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
  FrmRelSegurosPagos.MousePointer = vbDefault
  
End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelSegurosPagos
End Sub

Private Sub Form_Load()
  
  subManutencaoJanelasAtivas "I", "FrmRelSegurosPagos"
  
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
