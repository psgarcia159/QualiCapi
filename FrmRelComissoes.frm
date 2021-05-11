VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmRelComissoes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório Comissões dos Corretores"
   ClientHeight    =   3390
   ClientLeft      =   3330
   ClientTop       =   3615
   ClientWidth     =   6540
   Icon            =   "FrmRelComissoes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Período de Pamento"
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
      Left            =   1080
      TabIndex        =   14
      Top             =   1920
      Width           =   4650
      Begin MSComCtl2.DTPicker dtcInicioPagamento 
         Height          =   315
         Left            =   840
         TabIndex        =   15
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yy"
         DateIsNull      =   -1  'True
         Format          =   65208323
         CurrentDate     =   37692.3633680556
      End
      Begin MSComCtl2.DTPicker dtcFimPagamento 
         Height          =   315
         Left            =   3000
         TabIndex        =   16
         Top             =   285
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yy"
         DateIsNull      =   -1  'True
         Format          =   65208323
         CurrentDate     =   37692.3632291667
      End
      Begin VB.Label Label2 
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
         Left            =   2460
         TabIndex        =   18
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label1 
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
         Left            =   255
         TabIndex        =   17
         Top             =   330
         Width           =   495
      End
   End
   Begin VB.Frame FraPeriodo 
      Caption         =   "Período da Venda"
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
      Left            =   1080
      TabIndex        =   10
      Top             =   1140
      Width           =   4650
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   795
         TabIndex        =   5
         Top             =   285
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yy"
         DateIsNull      =   -1  'True
         Format          =   65208323
         CurrentDate     =   37692.3633680556
      End
      Begin MSComCtl2.DTPicker dtpFim 
         Height          =   315
         Left            =   3000
         TabIndex        =   6
         Top             =   285
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yy"
         DateIsNull      =   -1  'True
         Format          =   65208323
         CurrentDate     =   37692.3632291667
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
         Left            =   255
         TabIndex        =   12
         Top             =   330
         Width           =   495
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
         Left            =   2460
         TabIndex        =   11
         Top             =   330
         Width           =   495
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   5340
      TabIndex        =   8
      Top             =   2940
      Width           =   1035
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir..."
      Height          =   345
      Left            =   4200
      TabIndex        =   7
      Top             =   2940
      Width           =   1035
   End
   Begin MSDataListLib.DataCombo CboCorretores 
      Bindings        =   "FrmRelComissoes.frx":2AFA
      Height          =   315
      Left            =   1620
      TabIndex        =   3
      Top             =   660
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "corr_tx_Nome"
      BoundColumn     =   "corr_cd_Corretor"
      Text            =   "CboCorretores"
   End
   Begin MSAdodcLib.Adodc DatCorretores 
      Height          =   330
      Left            =   0
      Top             =   2940
      Visible         =   0   'False
      Width           =   1980
      _ExtentX        =   3493
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
      Caption         =   "DatCorretor"
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
   Begin Threed.SSCommand CmdLimpar 
      Height          =   315
      Left            =   6120
      TabIndex        =   4
      Top             =   660
      Width           =   315
      _Version        =   65536
      _ExtentX        =   556
      _ExtentY        =   556
      _StockProps     =   78
      MouseIcon       =   "FrmRelComissoes.frx":2B16
      Picture         =   "FrmRelComissoes.frx":2B32
   End
   Begin MSMask.MaskEdBox MskEmpreendimento 
      Height          =   315
      Left            =   1620
      TabIndex        =   0
      Top             =   255
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "AAAA"
      PromptChar      =   " "
   End
   Begin MSDataListLib.DataCombo CboEmpreendimento 
      Bindings        =   "FrmRelComissoes.frx":2C44
      Height          =   315
      Left            =   2205
      TabIndex        =   1
      Top             =   255
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "empd_tx_nome"
      BoundColumn     =   "empd_cd_empreendimento"
      Text            =   "CboEmpreendimento"
   End
   Begin Threed.SSCommand CmdLimparEmpreendimento 
      Height          =   315
      Left            =   6120
      TabIndex        =   2
      Top             =   240
      Width           =   315
      _Version        =   65536
      _ExtentX        =   556
      _ExtentY        =   556
      _StockProps     =   78
      MouseIcon       =   "FrmRelComissoes.frx":2C64
      Picture         =   "FrmRelComissoes.frx":2C80
   End
   Begin MSAdodcLib.Adodc DatEmpreendimento 
      Height          =   330
      Left            =   1980
      Top             =   2940
      Visible         =   0   'False
      Width           =   1980
      _ExtentX        =   3493
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
      Caption         =   "DatEmpreendimento"
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
   Begin VB.Label LblEmpreendimento 
      Alignment       =   1  'Right Justify
      Caption         =   "Empreendimento:"
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
      Left            =   120
      TabIndex        =   13
      Top             =   300
      Width           =   1485
   End
   Begin VB.Label LblCorretor 
      Alignment       =   1  'Right Justify
      Caption         =   "Corretor:"
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
      Left            =   840
      TabIndex        =   9
      Top             =   720
      Width           =   765
   End
End
Attribute VB_Name = "FrmRelComissoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CboCorretores_Click(Area As Integer)
  If CboCorretores.Text <> "" Then
        DatCorretores.Recordset.Bookmark = CboCorretores.SelectedItem
  End If
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
    Dim XLT_FILTROPERIODOPAGAMENTO As String
    Dim XLT_FILTROPERIODO As String 'Armazenar filtro do período escolhido
    Dim XLT_EMPREENDIMENTO As String 'Empreendimento escolhido
    Dim XLT_CORRETOR As String 'Corretor escolhido
    
    FrmRelComissoes.MousePointer = vbHourglass
 
    XGT_SELECAO = ""
    XGT_CONJUNCAO = ""
  
    'Período
    If (Not IsNull(DtpInicio.Value)) And (Not IsNull(DtpFim.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelComissoes.cont_dt_Venda >= " & FunNuloData(DtpInicio.Value, NomeSgbd) & "" & _
        " AND ConsCAPRelComissoes.cont_dt_Venda <= " & FunNuloData(DtpFim.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODO = "Período de Venda: " & Format$(DtpInicio.Value, "dd/mm/yy") & " a " & Format$(DtpFim.Value, "dd/mm/yy")
    ElseIf (Not IsNull(DtpInicio.Value)) And (IsNull(DtpFim.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelComissoes.cont_dt_Venda >= " & FunNuloData(DtpInicio.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODO = "Período de Venda: A partir de " & Format$(DtpInicio.Value, "dd/mm/yy")
    ElseIf (IsNull(DtpInicio.Value)) And (Not IsNull(DtpFim.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelComissoes.cont_dt_Venda <= " & FunNuloData(DtpFim.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODO = "Período de Venda: Até " & Format$(DtpFim.Value, "dd/mm/yy")
    Else
        XLT_FILTROPERIODO = "Período de Venda: Geral"
    End If
    
    'Período Pagamento
    If (Not IsNull(dtcInicioPagamento.Value)) And (Not IsNull(dtcFimPagamento.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelComissoes.comi_dt_Pagamento >= " & FunNuloData(dtcInicioPagamento.Value, NomeSgbd) & "" & _
        " AND ConsCAPRelComissoes.comi_dt_Pagamento <= " & FunNuloData(dtcFimPagamento.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODOPAGAMENTO = "Período de Pagamento: " & Format$(dtcInicioPagamento.Value, "dd/mm/yy") & " a " & Format$(DtpFim.Value, "dd/mm/yy")
    ElseIf (Not IsNull(dtcInicioPagamento.Value)) And (IsNull(dtcFimPagamento.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelComissoes.comi_dt_Pagamento >= " & FunNuloData(dtcInicioPagamento.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODOPAGAMENTO = "Período de Pagamento: A partir de " & Format$(dtcInicioPagamento.Value, "dd/mm/yy")
    ElseIf (IsNull(dtcInicioPagamento.Value)) And (Not IsNull(dtcFimPagamento.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelComissoes.comi_dt_Pagamento <= " & FunNuloData(dtcFimPagamento.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODOPAGAMENTO = "Período de Pagamento: Até " & Format$(dtcFimPagamento.Value, "dd/mm/yy")
    Else
        XLT_FILTROPERIODO = "Período de Pagamento: Geral"
    End If
    
    'Empreendimento
     If CboEmpreendimento.BoundText <> "" Then
       XLT_EMPREENDIMENTO = CboEmpreendimento.BoundText
       XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
       "ConsCAPRelComissoes.empd_cd_Empreendimento = '" & XLT_EMPREENDIMENTO & "'"
       XGT_CONJUNCAO = " AND "
     End If
     
     'Corretor
     If CboCorretores.BoundText <> "" Then
       XLT_CORRETOR = CboCorretores.BoundText
       XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
       "ConsCAPRelComissoes.corr_cd_Corretor = '" & XLT_CORRETOR & "'"
       XGT_CONJUNCAO = " AND "
     End If
     
             
    'Filtra a empresa Ativa
    XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + "ConsCAPRelComissoes.empr_cd_Empresa = " & CStr(PCodEmpresa)
  
    'Abre os recordsets e verifica se eles possuem dados
    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsCAPRelComissoes WHERE " & XGT_SELECAO)
    If XGO_RSRELATORIO.EOF Then
        MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
        FrmRelComissoes.MousePointer = vbDefault
        Exit Sub
    End If
    Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
  
    'Carrega os relatórios
    FrmVisRelatorios.Caption = "Relatório de Comissões dos Corretores"
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelComissoes.rpt")
  
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
  
    'Ativa os Relatórios
    With XGR_RELATORIO.FormulaFields
        .GetItemByName("Periodo").Text = Chr(34) & XLT_FILTROPERIODO & Chr(34)
        .GetItemByName("FiltroPeriodoPagamento").Text = Chr(34) & XLT_FILTROPERIODOPAGAMENTO & Chr(34)
        .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
    End With
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
    FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
    FrmRelComissoes.MousePointer = vbDefault
End Sub

Private Sub CmdLimpar_Click()
    CboCorretores.Text = ""
End Sub

Private Sub CmdLimparEmpreendimento_Click()
  CboEmpreendimento.Text = ""
  MskEmpreendimento.Text = ""
End Sub

Private Sub DTPicker2_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelComissoes
End Sub

Private Sub Form_Load()
  subManutencaoJanelasAtivas "I", "FrmRelComissoes"
  subConectarControleDadosNV DatCorretores, "SELECT * FROM Corretores ORDER BY corr_tx_Nome", Estatico
  subConectarControleDadosNV DatEmpreendimento, "SELECT * FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_tx_Nome", Estatico
End Sub

Private Sub MnuRelComissoes_Click()

End Sub

Private Sub MskEmpreendimento_GotFocus()
    Call subSelecionaMSK(MskEmpreendimento)
End Sub

Private Sub MskEmpreendimento_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
      KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub MskEmpreendimento_LostFocus()
    If MskEmpreendimento.Text <> "" Then
        SubPreencheComZero MskEmpreendimento
        CboEmpreendimento.BoundText = MskEmpreendimento.Text
        If CboEmpreendimento.Text = "" Then
            MsgBox "Não existe Empreendimento com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            MskEmpreendimento.SetFocus
        End If
    End If
End Sub
