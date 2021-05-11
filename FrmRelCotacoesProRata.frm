VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRelCotacoesProRata 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relação de Cotaçãoes"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   Icon            =   "FrmRelCotacoesProRata.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraProRata 
      Caption         =   "Correção Monetária"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   60
      TabIndex        =   8
      Top             =   1560
      Width           =   4095
      Begin VB.OptionButton OptCorrecProRataPosterior 
         Caption         =   "Pró-Rata Base Mês Posterior"
         Height          =   195
         Left            =   780
         TabIndex        =   10
         Top             =   540
         Width           =   2475
      End
      Begin VB.OptionButton OptCorrecProRataAnterior 
         Caption         =   "Pró-Rata Base Mês Anterior"
         Height          =   195
         Left            =   780
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   2835
      End
   End
   Begin VB.Frame FraIndexador 
      Height          =   675
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   4095
      Begin MSDataListLib.DataCombo CboIndexador 
         Bindings        =   "FrmRelCotacoesProRata.frx":2AFA
         Height          =   315
         Left            =   1665
         TabIndex        =   6
         Top             =   240
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "moed_tx_descricao"
         BoundColumn     =   "moed_cd_codmoeda"
         Text            =   "CboIndexador"
      End
      Begin VB.Label LblMoeda1 
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
         Left            =   480
         TabIndex        =   7
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir..."
      Height          =   345
      Left            =   1980
      TabIndex        =   4
      Top             =   2520
      Width           =   1035
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   3120
      TabIndex        =   3
      Top             =   2520
      Width           =   1035
   End
   Begin VB.Frame FraPeriodo 
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
      TabIndex        =   0
      Top             =   720
      Width           =   4110
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   1635
         TabIndex        =   1
         Top             =   300
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   20709379
         CurrentDate     =   38064.7902083333
      End
      Begin VB.Label LblInicio 
         Alignment       =   1  'Right Justify
         Caption         =   "Mês:"
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
         Left            =   1095
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
   End
   Begin MSAdodcLib.Adodc DatIndexador 
      Height          =   330
      Left            =   0
      Top             =   2460
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
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
      Caption         =   "DatMoeda1"
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
Attribute VB_Name = "FrmRelCotacoesProRata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancelar_Click()
   
   Unload Me
   
End Sub

'Private Sub CmdImprimir_Click()
'
'  Dim XLT_FILTROPERIODO As String 'Armazenar filtro do período escolhido
'  Dim XLT_FILTROINDEXADOR As String 'Indexador escolhido
'  Dim XLT_SQL As String
'
'  If CboIndexador.BoundText = "" Then
'    MsgBox "O indexador deve ser definido!", vbCritical, "ATENÇÃO"
'    Exit Sub
'  End If
'
'  Me.MousePointer = vbHourglass
'
'  XGT_SELECAO = ""
'  XGT_CONJUNCAO = ""
'
'  XLT_SQL = "SELECT *, " & _
'    "(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where " & _
'    "cotacoesmoedas.moed_cd_CodMoeda=" & CboIndexador.BoundText & _
'    "and cota_dt_datacotacao=" & FunNuloData("01/" & Month(dtpInicio.Value) & "/" & Year(dtpInicio.Value), NomeSgbd) & ") as Cotacao1 "
'
'  'Verifica se a pró rata é pelo mês anterior ou posterior
'  If OptCorrecProRataAnterior.Value = True Then
'    XLT_SQL = XLT_SQL & _
'      "(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where " & _
'      "cotacoesmoedas.moed_cd_CodMoeda=" & CboIndexador.BoundText & _
'      "and cota_dt_datacotacao= dateadd(mm,-1," & FunNuloData("01/" & Month(dtpInicio.Value) & "/" & Year(dtpInicio.Value), NomeSgbd) & ")) as Cotacao2 "
'  Else
'    XLT_SQL = XLT_SQL & _
'      "(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where " & _
'      "cotacoesmoedas.moed_cd_CodMoeda=" & CboIndexador.BoundText & _
'      "and cota_dt_datacotacao= dateadd(mm,1," & FunNuloData("01/" & Month(dtpInicio.Value) & "/" & Year(dtpInicio.Value), NomeSgbd) & ")) as Cotacao2 "
'  End If
'
'  XLT_SQL = XLT_SQL & " FROM ConsgenCotacoes"
'
'  'Indexador
'  XLT_FILTROINDEXADOR = "Indexador: " & CboIndexador.Text
'  XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
'  "moed_cd_CodMoeda = " & CboIndexador.BoundText
'  XGT_CONJUNCAO = " AND "
'
' 'Filtra a empresa Ativa
' 'XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + " empr_cd_Empresa = " & CStr(PCodEmpresa)
'
'  'WHERE
'  If XGT_SELECAO <> "" Then
'    XGT_SELECAO = " WHERE " & XGT_SELECAO
'  End If
'
'  'Abre os recordsets e verifica se eles possuem dados
'  Set XGO_RSRELATORIO = ConexaoRelatorio.Execute(XLT_SQL & XGT_SELECAO)
'  If XGO_RSRELATORIO.EOF Then
'      MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
'      Me.MousePointer = vbDefault
'      Exit Sub
'  End If
'  Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
'
'  'Carrega os relatórios
'  FrmVisRelatorios.Caption = "Relatório de Cotações PróRata"
'  Set RptCotacaoProRata = Nothing
'  Set XGR_RELATORIO = RptCotacaoProRata
'
'  'Passa o recordset para o relatório
'  XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
'
'  'Ativa os Relatórios
'  With XGR_RELATORIO.FormulaFields
'      .GetItemByName("Data1").Text = Chr(34) & Format(dtpInicio.Value, "01/mm/yy") & Chr(34)
'
'      If OptCorrecProRataAnterior.Value = True Then
'        .GetItemByName("Data2").Text = Chr(34) & DateAdd("m", -1, Format(dtpInicio.Value, "01/mm/yy")) & Chr(34)
'        .GetItemByName("ProRata").Text = Chr(34) & "A" & Chr(34)
'      Else
'        .GetItemByName("Data2").Text = Chr(34) & DateAdd("m", 1, Format(dtpInicio.Value, "01/mm/yy")) & Chr(34)
'        .GetItemByName("ProRata").Text = Chr(34) & "P" & Chr(34)
'      End If
'
'      .GetItemByName("FiltroPeriodo").Text = Chr(34) & "Mês: " & Format(dtpInicio.Value, "mm/yyyy") & Chr(34)
'      .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
'      .GetItemByName("FiltroIndexador").Text = Chr(34) & XLT_FILTROINDEXADOR & Chr(34)
'  End With
'  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
'  FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
'  Me.MousePointer = vbDefault
'
'End Sub
Private Sub CmdImprimir_Click()
  
  Dim XLT_FILTROPERIODO As String 'Armazenar filtro do período escolhido
  Dim XLT_FILTROINDEXADOR As String 'Indexador escolhido
  Dim XLT_SQL As String
  Dim XLD_DATA As Date
  Dim XLF_COTACAO As Double
  Dim XLT_PRORATA As String
  Dim XLO_COTACAO As New ADODB.Recordset
    
  If CboIndexador.BoundText = "" Then
    MsgBox "O indexador deve ser definido!", vbCritical, "ATENÇÃO"
    Exit Sub
  End If
  
  Me.MousePointer = vbHourglass

  XGT_SELECAO = ""
  XGT_CONJUNCAO = ""
                  
  XLT_SQL = "SELECT " & _
    "(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where " & _
    "cotacoesmoedas.moed_cd_CodMoeda=" & CboIndexador.BoundText & _
    "and cota_dt_datacotacao=" & FunNuloData("01/" & Month(dtpInicio.Value) & "/" & Year(dtpInicio.Value), NomeSgbd) & ") as Cotacao1, "

  'Verifica se a pró rata é pelo mês anterior ou posterior
  If OptCorrecProRataAnterior.Value = True Then
    XLT_PRORATA = "A"
    XLT_SQL = XLT_SQL & _
      "(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where " & _
      "cotacoesmoedas.moed_cd_CodMoeda=" & CboIndexador.BoundText & _
      "and cota_dt_datacotacao= dateadd(mm,-1," & FunNuloData("01/" & Month(dtpInicio.Value) & "/" & Year(dtpInicio.Value), NomeSgbd) & ")) as Cotacao2 "
  Else
    XLT_PRORATA = "P"
    XLT_SQL = XLT_SQL & _
      "(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where " & _
      "cotacoesmoedas.moed_cd_CodMoeda=" & CboIndexador.BoundText & _
      "and cota_dt_datacotacao= dateadd(mm,1," & FunNuloData("01/" & Month(dtpInicio.Value) & "/" & Year(dtpInicio.Value), NomeSgbd) & ")) as Cotacao2 "
  End If
  
  SubQOpenRecordset XLO_COTACAO, XLT_SQL, Estatico
  
  If IsNull(XLO_COTACAO!Cotacao1) Then
    MsgBox "Não existe cotação para a data " & Format(dtpInicio.Value, "\01/mm/yy"), vbCritical, "ATENÇÃO"
    Me.MousePointer = vbDefault
    Exit Sub
  ElseIf IsNull(XLO_COTACAO!Cotacao2) Then
    If XLT_PRORATA = "A" Then
      MsgBox "Não existe cotação para a data " & DateAdd("m", -1, Format(dtpInicio.Value, "\01/mm/yy")), vbCritical, "ATENÇÃO"
      Me.MousePointer = vbDefault
      Exit Sub
    Else
      MsgBox "Não existe cotação para a data " & DateAdd("m", 1, Format(dtpInicio.Value, "\01/mm/yy")), vbCritical, "ATENÇÃO"
      Me.MousePointer = vbDefault
      Exit Sub
    End If
  End If
                  
  ConexaoRelatorio.Execute "CREATE TABLE #Cotacao_Temp " & _
                  " ( Data datetime," & _
                  " Cotacao decimal(9, 4))"
  
  XLD_DATA = Format(dtpInicio.Value, "\01/mm/yy")
                  
  While XLD_DATA < DateAdd("m", 1, Format(dtpInicio.Value, "\01/mm/yy"))
    
    XLF_COTACAO = FunCorrecaoMonetaria(1, XLO_COTACAO!Cotacao1, XLO_COTACAO!Cotacao2, "M", XLT_PRORATA, XLD_DATA, XLD_DATA, CboIndexador.BoundText)
    'XLF_COTACAO = FunCorrecaoMonetaria(1, XLO_COTACAO!Cotacao1, XLO_COTACAO!Cotacao2, "M", XLT_PRORATA, XLD_DATA)
    
    ConexaoRelatorio.Execute "INSERT INTO #Cotacao_Temp (Data,Cotacao) VALUES (" & FunNuloData(XLD_DATA, NomeSgbd) & "," & FunTrataFloat(XLF_COTACAO) & ")"
    
    XLD_DATA = DateAdd("d", 1, XLD_DATA)
  Wend
  
  'Abre os recordsets e verifica se eles possuem dados
  Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM #Cotacao_Temp")
  If XGO_RSRELATORIO.EOF Then
      MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
      Me.MousePointer = vbDefault
      Exit Sub
  End If
  Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)

  'Carrega os relatórios
  FrmVisRelatorios.Caption = "Relatório de Cotações PróRata"
  Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelCotacoesProrata.rpt")

  'Passa o recordset para o relatório
  XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
  
  'Apaga a tabela temporária
  ConexaoRelatorio.Execute "drop table #Cotacao_Temp"

  'Ativa os Relatórios
  With XGR_RELATORIO.FormulaFields
      .GetItemByName("FiltroPeriodo").Text = Chr(34) & "Mês: " & Format(dtpInicio.Value, "mm/yyyy") & Chr(34)
      .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
      .GetItemByName("FiltroIndexador").Text = Chr(34) & "Indexador: " & CboIndexador.Text & Chr(34)
      If XLT_PRORATA = "A" Then
        .GetItemByName("ProRata").Text = Chr(34) & "Pro-Rata: Mês Anterior" & Chr(34)
      Else
        .GetItemByName("ProRata").Text = Chr(34) & "Pro-Rata: Mês Posterior" & Chr(34)
      End If
  End With
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
  FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
  Me.MousePointer = vbDefault
  
End Sub

Private Sub Form_Load()
  
  subConectarControleDadosNV DatIndexador, "SELECT * FROM Moedas WHERE moed_tx_Indexacao='M' ORDER BY moed_tx_descricao", Estatico
  
  If Not DatIndexador.Recordset.EOF Then
    CboIndexador.BoundText = DatIndexador.Recordset!moed_cd_CodMoeda
  End If
  
End Sub
