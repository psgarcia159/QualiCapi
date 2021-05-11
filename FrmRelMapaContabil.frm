VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRelMapaContabil 
   Caption         =   "Mapa Contábil de Clientes"
   ClientHeight    =   3045
   ClientLeft      =   945
   ClientTop       =   2190
   ClientWidth     =   5160
   Icon            =   "FrmRelMapaContabil.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3045
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2595
      Left            =   0
      TabIndex        =   8
      Top             =   -60
      Width           =   5115
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
         Height          =   1035
         Left            =   60
         TabIndex        =   11
         Top             =   1500
         Width           =   4995
         Begin VB.OptionButton OptValorHistoricoSocietario 
            Caption         =   "Valor Histórico Societário"
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
            Left            =   180
            TabIndex        =   15
            Top             =   600
            Visible         =   0   'False
            Width           =   3015
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
            Left            =   3480
            TabIndex        =   5
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
            Left            =   1920
            TabIndex        =   4
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
            Left            =   180
            TabIndex        =   3
            Top             =   300
            Width           =   1575
         End
      End
      Begin VB.Frame FraModelo 
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
         TabIndex        =   10
         Top             =   780
         Width           =   4995
         Begin MSComCtl2.DTPicker DtpDtBase 
            Height          =   315
            Left            =   2580
            TabIndex        =   2
            Top             =   240
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MM/yyyy"
            Format          =   99418115
            CurrentDate     =   37636
         End
         Begin VB.Label LblInicial 
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
            Height          =   255
            Left            =   780
            TabIndex        =   12
            Top             =   300
            Width           =   1815
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
         TabIndex        =   9
         Top             =   120
         Width           =   4995
         Begin MSMask.MaskEdBox MskEmpreendimentoInicial 
            Height          =   315
            Left            =   2400
            TabIndex        =   0
            Top             =   180
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "AAAA"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MskEmpreendimentoFinal 
            Height          =   315
            Left            =   3375
            TabIndex        =   1
            Top             =   195
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "AAAA"
            PromptChar      =   " "
         End
         Begin VB.Label LblEmpreedimento 
            Caption         =   "Empreedimento:"
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
            Left            =   960
            TabIndex        =   14
            Top             =   240
            Width           =   1455
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
            Left            =   3075
            TabIndex        =   13
            Top             =   255
            Width           =   195
         End
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir..."
      Height          =   345
      Left            =   2760
      TabIndex        =   6
      Top             =   2640
      Width           =   1035
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3960
      TabIndex        =   7
      Top             =   2640
      Width           =   1035
   End
End
Attribute VB_Name = "FrmRelMapaContabil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
    
  Dim XLT_FILTROOBSERVACAO As String
  Dim XLT_INDEXADOR As String
  Dim XLT_MODELO As String
  Dim XLT_SQL As String
  Dim XLT_FILTROEMPREENDIMENTOS As String
  Dim XLD_DATABASE As Date
  
  Me.MousePointer = vbHourglass
  
  XGT_SELECAO = ""
  XGT_CONJUNCAO = ""
  
  XGT_SELECAO = ""
  XGT_CONJUNCAO = ""
  
  XLD_DATABASE = Format(DtpDtBase.Value, "\01/mm/yy")
  
 'Cotação do mês atual
   XLT_SQL = "" & _
   "SELECT ConsCAPRelExtratoContrato.*," & _
   "CASE Indexacao1   when 'M' then" & _
     " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(XLD_DATABASE, NomeSgbd) & ")" & _
     " when 'D' then" & _
     " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(XLD_DATABASE, NomeSgbd) & ")" & _
   " end as Cotacao1," & _
   "CASE  Indexacao2 when 'M' then" & _
     " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(XLD_DATABASE, NomeSgbd) & ")" & _
     " when 'D'  then" & _
     " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(XLD_DATABASE, NomeSgbd) & ")" & _
   " end as Cotacao2,"

   'Cotação do mês pagamento se for mensal é no dia 1o
   XLT_SQL = XLT_SQL & _
    "CASE  Indexacao3 when 'M' then" & _
     " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=convert(varchar,month(titu_dt_BasePagto))+'/01/' + convert(varchar,year(titu_dt_BasePagto)))" & _
     " when 'D'  then" & _
     " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=titu_dt_BasePagto)" & _
   " end as Cotacao3,"
   
   'Cotação do mês atual -1
   XLT_SQL = XLT_SQL & _
   "CASE Indexacao1   when 'M' then" & _
     " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(DateAdd("m", -1, XLD_DATABASE), NomeSgbd) & ")" & _
     " when 'D' then" & _
     " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(DateAdd("m", -1, XLD_DATABASE), NomeSgbd) & ")" & _
   " end as Cotacao4," & _
   "CASE  Indexacao2 when 'M' then" & _
     " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(DateAdd("m", -1, XLD_DATABASE), NomeSgbd) & ")" & _
     " when 'D'  then" & _
     " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(DateAdd("m", -1, XLD_DATABASE), NomeSgbd) & ")" & _
   " end as Cotacao5,"
  
  'Cotação do mês anterior ao mês atual ou do mês posterior ao mês atual dependendo da correcao pro-rata
  XLT_SQL = XLT_SQL & _
   "CASE  Indexacao1 when 'M' then " & _
     "CASE cont_tx_CorrecProrata when 'A' then" & _
       " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData(XLD_DATABASE, NomeSgbd) & "))" & _
     "when 'P' then" & _
       " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,1," & FunNuloData(XLD_DATABASE, NomeSgbd) & "))" & _
     "end " & _
   " when 'D' then Null " & _
   " end as Cotacao1_MesAnterior,"

   XLT_SQL = XLT_SQL & _
   "CASE Indexacao2 when 'M' then " & _
     "CASE cont_tx_CorrecProrata when 'A' then" & _
       " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData(XLD_DATABASE, NomeSgbd) & "))" & _
     "when 'P' then" & _
       " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,1," & FunNuloData(XLD_DATABASE, NomeSgbd) & "))" & _
     "end " & _
   " when 'D' then Null " & _
   " end As Cotacao2_MesAnterior,"
    
   'Cotação do mês anterior ao pagamento ou do mês posterior ao pagamento
   'dependendo da correcao pro-rata, se for mesnal é dia 1o
   XLT_SQL = XLT_SQL & _
   "CASE Indexacao3 when 'M' then " & _
     "CASE cont_tx_CorrecProrata when 'A' then" & _
       " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1,convert(varchar,month(titu_dt_BasePagto))+'/01/' + convert(varchar,year(titu_dt_BasePagto))))" & _
     "when 'P' then" & _
       " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda3=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,1,convert(varchar,month(titu_dt_BasePagto))+'/01/' + convert(varchar,year(titu_dt_BasePagto))))" & _
     "end " & _
   " when 'D' then Null " & _
   " end As Cotacao3_MesAnterior,"
   
  'Cotação do mês anterior ao mês atual-1 ou do mês posterior ao mês atual-1 dependendo da correcao pro-rata
  XLT_SQL = XLT_SQL & _
   "CASE  Indexacao1 when 'M' then " & _
     "CASE cont_tx_CorrecProrata when 'A' then" & _
       " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData(DateAdd("m", -1, XLD_DATABASE), NomeSgbd) & "))" & _
     "when 'P' then" & _
       " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,1," & FunNuloData(DateAdd("m", -1, XLD_DATABASE), NomeSgbd) & "))" & _
     "end " & _
   " when 'D' then Null " & _
   " end as Cotacao4_MesAnterior,"

   XLT_SQL = XLT_SQL & _
   "CASE Indexacao2 when 'M' then " & _
     "CASE cont_tx_CorrecProrata when 'A' then" & _
       " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData(DateAdd("m", -1, XLD_DATABASE), NomeSgbd) & "))" & _
     "when 'P' then" & _
       " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,1," & FunNuloData(DateAdd("m", -1, XLD_DATABASE), NomeSgbd) & "))" & _
     "end " & _
   " when 'D' then Null " & _
   " end As Cotacao5_MesAnterior" & _
   " From ConsCAPRelExtratoContrato WHERE " '21/12/09
  
  'Status
  XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
    "(ConsCAPRelExtratoContrato.cont_dt_Venda <= " & FunNuloData(Format(DateAdd("d", -1, DateAdd("m", 1, XLD_DATABASE))), NomeSgbd) & _
    " AND (ConsCAPRelExtratoContrato.cont_dt_Distrato is null OR " & _
    " (ConsCAPRelExtratoContrato.cont_dt_Distrato >= " & FunNuloData(Format(DtpDtBase.Value, "\01/mm/yy"), NomeSgbd) & " AND " & _
        " ConsCAPRelExtratoContrato.cont_tx_Status<>'C'))) "
  XGT_CONJUNCAO = " AND "
    
  'Empreendimento
  If MskEmpreendimentoInicial.Text <> "" And MskEmpreendimentoFinal.Text <> "" Then
    XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
    "consCAPRelextratoContrato.empd_cd_Empreendimento >= '" & MskEmpreendimentoInicial.Text & "'" & _
    " AND consCAPRelextratoContrato.empd_cd_Empreendimento <= '" & MskEmpreendimentoFinal.Text & "'"
    XGT_CONJUNCAO = " AND "
    XLT_FILTROEMPREENDIMENTOS = "Empreendimento: " & MskEmpreendimentoInicial & " a " & MskEmpreendimentoFinal
  ElseIf MskEmpreendimentoInicial.Text <> "" And MskEmpreendimentoFinal.Text = "" Then
    XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
    "consCAPRelextratoContrato.empd_cd_Emprendimento >= '" & MskEmpreendimentoInicial.Text & "'"
    XGT_CONJUNCAO = " AND "
    XLT_FILTROEMPREENDIMENTOS = "Empreendimento: A partir de " & MskEmpreendimentoInicial
  ElseIf MskEmpreendimentoInicial.Text = "" And MskEmpreendimentoFinal.Text <> "" Then
    XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
    "consCAPRelextratoContrato.empd_cd_Emprendimento <= '" & MskEmpreendimentoFinal.Text & "'"
    XGT_CONJUNCAO = " AND "
    XLT_FILTROEMPREENDIMENTOS = "Empreendimento: Até " & MskEmpreendimentoInicial
  Else
    XLT_FILTROEMPREENDIMENTOS = "Empreendimento: Geral"
  End If
  
  'XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
   ' "consCAPRelextratoContrato.Imovel='0530.1001'"
  
  'Indexador
  If (OptValorHistorico.Value = True) Or (OptValorHistoricoSocietario.Value = True) Then
    XLT_INDEXADOR = "0"
  ElseIf OptIndexador1.Value = True Then
    XLT_INDEXADOR = "1"
  ElseIf OptIndexador2.Value = True Then
    XLT_INDEXADOR = "2"
  End If
  
  'Pagos só depois do mês
  XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
    " (ConsCAPRelExtratoContrato.titu_dt_Pagamento is null OR" & _
    " ConsCAPRelExtratoContrato.titu_dt_Pagamento >= " & FunNuloData(XLD_DATABASE, NomeSgbd) & ")"
    
  'Filtra a empresa Ativa
  XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + "consCAPRelextratoContrato.empr_cd_Empresa = " & CStr(PCodEmpresa)

  'Ordenação
  'XGT_SELECAO = XGT_SELECAO + " ORDER BY empd_cd_Empreendimento"
  
  'Fecha o recordset se estiver aberto
  If XGO_RSRELATORIO.State = adStateOpen Then
    XGO_RSRELATORIO.Close
    Set XGO_RSRELATORIO = Nothing
  End If
  'Abre os recordsets e verifica se eles possuem dados
  Set XGO_RSRELATORIO = ConexaoRelatorio.Execute(XLT_SQL & XGT_SELECAO)
  
  If XGO_RSRELATORIO.EOF Then
      MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
      Me.MousePointer = vbDefault
      Exit Sub
  End If
  
  'Fecha o recordset se estiver aberto
  If XGO_RSCABECALHO.State = adStateOpen Then
    XGO_RSCABECALHO.Close
    Set XGO_RSCABECALHO = Nothing
  End If
  Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
    
  Set XGR_RELATORIO = Nothing
  If OptValorHistoricoSocietario.Value = True Then
    Set XGR_RELATORIO = RptMapaContabil_Societario
  Else
    Set XGR_RELATORIO = RptMapaContabil
  End If
  
  XGR_RELATORIO.DiscardSavedData
  
  'Passa o recordset para o relatório
  XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO

  'Ativa os Relatórios
  With XGR_RELATORIO.FormulaFields
      .GetItemByName("FiltroDataBase").Text = Chr(34) & Format(DtpDtBase.Value, "\01/mm/yy") & Chr(34)
      .GetItemByName("Indexador").Text = Chr(34) & XLT_INDEXADOR & Chr(34)
      .GetItemByName("FiltroEmpreendimento").Text = Chr(34) & XLT_FILTROEMPREENDIMENTOS & Chr(34)
      .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
  End With
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
  FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
  Me.MousePointer = vbDefault
End Sub


Private Sub Form_Activate()
  Set Formulario = FrmRelMapaContabil
End Sub

Private Sub Form_Load()
  
  subManutencaoJanelasAtivas "I", "FrmRelMapaContabil"
  'subConectarControleDadosNV DatObs, "SELECT * FROM Observacoes ORDER BY obse_tx_observacao", Estatico
  DtpDtBase.Value = Date
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  If XGO_RSRELATORIO.State = adStateOpen Then
    XGO_RSRELATORIO.Close
  End If
  
  If XGO_RSCABECALHO.State = adStateOpen Then
    XGO_RSCABECALHO.Close
  End If
  
  Set XGO_RSRELATORIO = Nothing
  Set XGO_RSCABECALHO = Nothing
  
End Sub

Private Sub OptIndexador1_Click()
OptValorHistoricoSocietario.Visible = False '22/12/09
End Sub

Private Sub OptIndexador2_Click()
OptValorHistoricoSocietario.Visible = False '22/12/09
End Sub

Private Sub OptValorHistorico_Click()
OptValorHistoricoSocietario.Visible = True '22/12/09
End Sub
