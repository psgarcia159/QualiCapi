VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRelCOAF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transações Imobiliárias"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   LinkTopic       =   "FrmRelCOAF"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   6810
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   5640
      TabIndex        =   6
      Top             =   1560
      Width           =   1035
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir..."
      Height          =   345
      Left            =   4440
      TabIndex        =   5
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Frame FrmPeriodo 
      Caption         =   "Período da Venda:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin MSComCtl2.DTPicker DtpInicio 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yy"
         Format          =   68943875
         CurrentDate     =   39052.7344560185
      End
      Begin MSComCtl2.DTPicker DtpFim 
         Height          =   315
         Left            =   4200
         TabIndex        =   2
         Top             =   255
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yy"
         Format          =   68943875
         CurrentDate     =   39082.7346064815
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
         Left            =   3660
         TabIndex        =   4
         Top             =   315
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
         Left            =   900
         TabIndex        =   3
         Top             =   315
         Width           =   495
      End
   End
   Begin MSAdodcLib.Adodc DatEmpreendimento 
      Height          =   330
      Left            =   120
      Top             =   1440
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
   Begin MSMask.MaskEdBox MskEmpreendimento 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Top             =   1080
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
      Bindings        =   "FrmRelCOAF.frx":0000
      Height          =   315
      Left            =   2280
      TabIndex        =   8
      Top             =   1080
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "empd_tx_nome"
      BoundColumn     =   "empd_cd_empreendimento"
      Text            =   "CboEmpreendimento"
   End
   Begin Threed.SSCommand cmdLimparEmpreendimento 
      Height          =   315
      Left            =   6360
      TabIndex        =   9
      Top             =   1080
      Width           =   315
      _Version        =   65536
      _ExtentX        =   556
      _ExtentY        =   556
      _StockProps     =   78
      MouseIcon       =   "FrmRelCOAF.frx":0020
      Picture         =   "FrmRelCOAF.frx":003C
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
      TabIndex        =   10
      Top             =   1080
      Width           =   1485
   End
End
Attribute VB_Name = "FrmRelCOAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CboEmpreendimento_Change()

    If CboEmpreendimento.Text <> "" Then
      DatEmpreendimento.Recordset.bookmark = CboEmpreendimento.SelectedItem
      MskEmpreendimento.Text = CboEmpreendimento.BoundText
''      subConectarControleDadosNV DatClientes, "SELECT focl_cd_FornCli,focl_tx_RazaoSocial FROM ConsCAPImoveisPorCliente " & _
''        " WHERE empr_cd_Empresa= " & PCodEmpresa & " AND empd_cd_Empreendimento='" & CboEmpreendimento.BoundText & "'" & _
''        " ORDER BY focl_tx_RazaoSocial", Estatico
'    Else
'      subConectarControleDadosNV DatClientes, "SELECT focl_cd_FornCli,focl_tx_RazaoSocial " & _
'      "FROM ConsCAPImoveisPorCliente WHERE empr_cd_Empresa= " & PCodEmpresa & " ORDER BY focl_tx_RazaoSocial", Estatico
    End If
'    CboCliente.BoundText = ""
'
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
    Dim XLT_FILTROPERIODO As String 'Armazenar filtro do período escolhido
    Dim XLT_EMPREENDIMENTO As String 'Empreendimento escolhido
    Dim XFO_TRANSACAOPF As New ADODB.Recordset
    Dim XFO_TRANSACAOPJ As New ADODB.Recordset
    
    FrmRelCOAF.MousePointer = vbHourglass
 
    XGT_SELECAO = ""
    XGT_CONJUNCAO = ""
                  
    'Empreendimento
    If CboEmpreendimento.BoundText <> "" Then
      XLT_EMPREENDIMENTO = CboEmpreendimento.BoundText
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "empd_cd_Empreendimento = '" & XLT_EMPREENDIMENTO & "'"
      XGT_CONJUNCAO = " AND "
    End If
    
    'Data da Venda
    If (Not IsNull(DtpInicio.Value)) And (Not IsNull(DtpFim.Value)) Then
       XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "cont_dt_Venda >= " & FunNuloData(DtpInicio.Value, NomeSgbd) & "" & _
        " AND cont_dt_Venda <= " & FunNuloData(DtpFim.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODO = "Período de Venda: " & Format$(DtpInicio.Value, "dd/mm/yy") & " a " & Format$(DtpFim.Value, "dd/mm/yy")
    ElseIf (Not IsNull(DtpInicio.Value)) And (IsNull(DtpFim.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "cont_dt_Venda >= " & FunNuloData(DtpInicio.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODO = "Período de Venda: A partir de " & Format$(DtpInicio.Value, "dd/mm/yy")
    ElseIf (IsNull(DtpInicio.Value)) And (Not IsNull(DtpFim.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "(cont_dt_Venda <= " & FunNuloData(DtpFim.Value, NomeSgbd) & ")"
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODO = "Período de Venda: Até " & Format$(DtpFim.Value, "dd/mm/yy")
    Else
        XLT_FILTROPERIODO = "Período de Venda: Geral"
    End If

    'Filtra a empresa Ativa
    XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + "ConsCAPIRelTransacoesImobiliarias.empr_cd_empresa = " & CStr(PCodEmpresa)
    
    'XGT_SELECAO = XGT_SELECAO + " ORDER BY focl_tx_razaosocial"
    
    'Fecha o recordset se estiver aberto
    If XGO_RSRELATORIO.State = adStateOpen Then
      XGO_RSRELATORIO.Close
      Set XGO_RSRELATORIO = Nothing
    End If
    
    'Abre os recordsets e verifica se eles possuem dados
    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)

    If XGO_RSRELATORIO.EOF Then
        MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
        FrmRelItiv.MousePointer = vbDefault
         FrmRelCOAF.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Fecha o recordset se estiver aberto
    If XGO_RSCABECALHO.State = adStateOpen Then
      XGO_RSCABECALHO.Close
      Set XGO_RSCABECALHO = Nothing
    End If
    Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)

    'Carrega os relatórios
    FrmVisRelatorios.Caption = "Relatório de Transações Imobiliárias"
   
    Set XGR_RELATORIO = Nothing
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelTransacoesImobiliarias.rpt")

    'Fecha o recordset se estiver aberto
    If XFO_TRANSACAOPF.State = adStateOpen Then
        XFO_TRANSACAOPF.Close
        Set XFO_TRANSACAOPF = Nothing
    End If
    
    Set XFO_TRANSACAOPF = ConexaoRelatorio.Execute("SELECT * FROM ConsCAPIRelTransacoesImobiliarias WHERE " & XGT_SELECAO & " AND focl_tx_tipo = 'F' ORDER BY focl_tx_razaosocial")
    
    'Fecha o recordset se estiver aberto
    If XFO_TRANSACAOPJ.State = adStateOpen Then
        XFO_TRANSACAOPJ.Close
        Set XFO_TRANSACAOPJ = Nothing
    End If
    
    Set XFO_TRANSACAOPJ = ConexaoRelatorio.Execute("SELECT * FROM ConsCAPIRelTransacoesImobiliarias WHERE " & XGT_SELECAO & " AND focl_tx_tipo = 'J' ORDER BY focl_tx_razaosocial")
    
      
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
  
    'Ativa os Relatórios
    With XGR_RELATORIO.FormulaFields
        .GetItemByName("Periodo").Text = Chr(34) & XLT_FILTROPERIODO & Chr(34)
       '.GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
    End With
    
    'FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XFO_TRANSACAOPJ, "RelTransacoesPJ"
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XFO_TRANSACAOPF, "RelTransacoesPF"

    FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
        
    FrmRelCOAF.MousePointer = vbDefault
    
End Sub

Private Sub CmdLimpar_Click()
    CboCliente.Text = ""
End Sub

Private Sub CmdLimparEmpreendimento_Click()
    CboEmpreendimento.BoundText = ""
    MskEmpreendimento.Text = ""
End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelCOAF
End Sub

Private Sub Form_Load()
  
  subManutencaoJanelasAtivas "I", "FrmRelCOAF"
  
  'subConectarControleDadosNV DatClientes, "SELECT focl_cd_forncli,focl_tx_classe,focl_tx_razaoSocial FROM FornClientes WHERE focl_tx_classe = 'C' ORDER BY focl_tx_razaoSocial", Estatico
  'subConectarControleDadosNV DatClientes, "SELECT focl_cd_FornCli,focl_tx_RazaoSocial " & _
  '"FROM ConsCAPImoveisPorCliente WHERE empr_cd_Empresa= " & PCodEmpresa & " ORDER BY focl_tx_RazaoSocial", Estatico
  
  subConectarControleDadosNV DatEmpreendimento, "SELECT * FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_tx_Nome", Estatico

End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  If XGO_RSRELATORIO.State = adStateOpen Then
    XGO_RSRELATORIO.Close
  End If
  
  If XGO_RSCABECALHO.State = adStateOpen Then
    XGO_RSCABECALHO.Close
  End If
 
'  If XFO_TRANSACAOPF.State = adStateOpen Then
'    XFO_TRANSACAOPF.Close
'  End If
'
'  If XFO_TRANSACAOPJ.State = adStateOpen Then
'    XFO_TRANSACAOPJ.Close
'  End If
 
  Set XGO_RSRELATORIO = Nothing
  Set XGO_RSCABECALHO = Nothing
  Set XFO_TRANSACAOPF = Nothing
  Set XFO_TRANSACAOPJ = Nothing
  
End Sub


Private Sub MskEmpreendimento_LostFocus()
  
  If MskEmpreendimento.Text <> "" Then
    SubPreencheComZero MskEmpreendimento
    CboEmpreendimento.BoundText = MskEmpreendimento.Text
    If CboEmpreendimento.Text = "" Then
      MsgBox "Não existe Empreendimento com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
      MskEmpreendimento.SetFocus
      Exit Sub
    End If
  Else
    CboEmpreendimento.BoundText = ""
  End If
  
End Sub

Private Sub MskEmpreendimento_GotFocus()
    Call subSelecionaMSK(MskEmpreendimento)
End Sub

Private Sub MskEmpreendimento_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
      KeyAscii = KeyAscii - 32
    End If
End Sub




