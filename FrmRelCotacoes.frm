VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRelCotacoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relação de Cotaçãoes"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   Icon            =   "FrmRelCotacoes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraIndexador 
      Height          =   675
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   4095
      Begin MSDataListLib.DataCombo CboIndexador 
         Bindings        =   "FrmRelCotacoes.frx":2AFA
         Height          =   315
         Left            =   1425
         TabIndex        =   8
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
      Begin Threed.SSCommand CmdLimparIndex1 
         Height          =   315
         Left            =   3300
         TabIndex        =   9
         Top             =   240
         Width           =   360
         _Version        =   65536
         _ExtentX        =   635
         _ExtentY        =   556
         _StockProps     =   78
         MouseIcon       =   "FrmRelCotacoes.frx":2B15
         Picture         =   "FrmRelCotacoes.frx":2F67
      End
      Begin VB.Label LblMoeda1 
         Alignment       =   1  'Right Justify
         Caption         =   "Indexador 1:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir..."
      Height          =   345
      Left            =   1980
      TabIndex        =   6
      Top             =   1620
      Width           =   1035
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   3120
      TabIndex        =   5
      Top             =   1620
      Width           =   1035
   End
   Begin VB.Frame FraPeriodo 
      Caption         =   "Período"
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
         Left            =   675
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yy"
         DateIsNull      =   -1  'True
         Format          =   3932163
         CurrentDate     =   37692.3633680556
      End
      Begin MSComCtl2.DTPicker dtpFim 
         Height          =   315
         Left            =   2640
         TabIndex        =   2
         Top             =   285
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yy"
         DateIsNull      =   -1  'True
         Format          =   3932163
         CurrentDate     =   37692.3632291667
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
         Left            =   2100
         TabIndex        =   4
         Top             =   330
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
         Left            =   135
         TabIndex        =   3
         Top             =   330
         Width           =   495
      End
   End
   Begin MSAdodcLib.Adodc DatIndexador 
      Height          =   330
      Left            =   0
      Top             =   1560
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
Attribute VB_Name = "FrmRelCotacoes"
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
  Dim XLT_FILTROINDEXADOR As String 'Indexador escolhido
    
  Me.MousePointer = vbHourglass

  XGT_SELECAO = ""
  XGT_CONJUNCAO = ""

  'Período
  If (Not IsNull(dtpInicio.Value)) And (Not IsNull(dtpFim.Value)) Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "cota_dt_DataCotacao >= " & FunNuloData(dtpInicio.Value, NomeSgbd) & "" & _
      " AND cota_dt_DataCotacao <= " & FunNuloData(dtpFim.Value, NomeSgbd) & ""
      XGT_CONJUNCAO = " AND "
      XLT_FILTROPERIODO = "Período: " & Format$(dtpInicio.Value, "dd/mm/yy") & " a " & Format$(dtpFim.Value, "dd/mm/yy")
  ElseIf (Not IsNull(dtpInicio.Value)) And (IsNull(dtpFim.Value)) Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "cota_dt_DataCotacao >= " & FunNuloData(dtpInicio.Value, NomeSgbd) & ""
      XGT_CONJUNCAO = " AND "
      XLT_FILTROPERIODO = "Período: A partir de " & Format$(dtpInicio.Value, "dd/mm/yy")
  ElseIf (IsNull(dtpInicio.Value)) And (Not IsNull(dtpFim.Value)) Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "cota_dt_DataCotacao <= " & FunNuloData(dtpFim.Value, NomeSgbd) & ""
      XGT_CONJUNCAO = " AND "
      XLT_FILTROPERIODO = "Período: Até " & Format$(dtpFim.Value, "dd/mm/yy")
  Else
      XLT_FILTROPERIODO = "Período: Geral"
  End If
  
  'Indexador
   If CboIndexador.BoundText <> "" Then
     XLT_FILTROINDEXADOR = "Indexador: " & CboIndexador.Text
     XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
     "moed_cd_CodMoeda = " & CboIndexador.BoundText
     XGT_CONJUNCAO = " AND "
   Else
     XLT_FILTROINDEXADOR = "Indexador: Geral"
   End If
           
  'Filtra a empresa Ativa
  'XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + " empr_cd_Empresa = " & CStr(PCodEmpresa)
  
  'WHERE
  If XGT_SELECAO <> "" Then
    XGT_SELECAO = " WHERE " & XGT_SELECAO
  End If

  'Abre os recordsets e verifica se eles possuem dados
  Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsGENCotacoes " & XGT_SELECAO & "ORDER BY moed_tx_descricao,cota_dt_DataCotacao")
  If XGO_RSRELATORIO.EOF Then
      MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
      Me.MousePointer = vbDefault
      Exit Sub
  End If
  Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)

  'Carrega os relatórios
  FrmVisRelatorios.Caption = "Relatório de Cotaçãoes"
  Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelCotacoes.rpt")

  'Passa o recordset para o relatório
  XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO

  'Ativa os Relatórios
  With XGR_RELATORIO.FormulaFields
      .GetItemByName("FiltroPeriodo").Text = Chr(34) & XLT_FILTROPERIODO & Chr(34)
      .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
      .GetItemByName("FiltroIndexador").Text = Chr(34) & XLT_FILTROINDEXADOR & Chr(34)
  End With
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
  FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
  Me.MousePointer = vbDefault
  
End Sub

Private Sub Form_Load()
  
  subConectarControleDadosNV DatIndexador, "SELECT * FROM Moedas ORDER BY moed_tx_descricao", Estatico
  
End Sub
