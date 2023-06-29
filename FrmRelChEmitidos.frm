VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRelChEmitidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Cheques Emitidos"
   ClientHeight    =   1935
   ClientLeft      =   2640
   ClientTop       =   3615
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmRelChEmitidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6375
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir..."
      Height          =   375
      Left            =   4020
      TabIndex        =   3
      Top             =   1500
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5220
      TabIndex        =   4
      Top             =   1500
      Width           =   1095
   End
   Begin VB.Frame FraPeriodo 
      Caption         =   "Período (Emissão)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   870
      TabIndex        =   5
      Top             =   105
      Width           =   4605
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   720
         TabIndex        =   0
         Top             =   300
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   95420417
         CurrentDate     =   37456
      End
      Begin MSComCtl2.DTPicker dtpFim 
         Height          =   315
         Left            =   2880
         TabIndex        =   1
         Top             =   300
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   95420417
         CurrentDate     =   37456
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Início:"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   345
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Fim:"
         Height          =   180
         Left            =   2460
         TabIndex        =   6
         Top             =   345
         Width           =   375
      End
   End
   Begin MSAdodcLib.Adodc DatBanco 
      Height          =   330
      Left            =   120
      Top             =   1440
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "DatBanco"
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
   Begin MSDataListLib.DataCombo CboCCorrente2 
      Bindings        =   "FrmRelChEmitidos.frx":2AFA
      Height          =   330
      Left            =   810
      TabIndex        =   2
      Top             =   1035
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "desccombo"
      BoundColumn     =   "coco_cd_codigo"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand CmdLimpar 
      Height          =   330
      Left            =   5960
      TabIndex        =   9
      Top             =   1020
      Width           =   345
      _Version        =   65536
      _ExtentX        =   609
      _ExtentY        =   582
      _StockProps     =   78
      MouseIcon       =   "FrmRelChEmitidos.frx":2B11
      Picture         =   "FrmRelChEmitidos.frx":2F63
   End
   Begin VB.Label LblBanco 
      Alignment       =   1  'Right Justify
      Caption         =   "C/C:"
      Height          =   210
      Left            =   135
      TabIndex        =   8
      Top             =   1110
      Width           =   615
   End
End
Attribute VB_Name = "FrmRelChEmitidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CboCCorrente2_Change()
    If CboCCorrente2.Text <> "" Then
        DatBanco.Recordset.bookmark = CboCCorrente2.SelectedItem
    End If
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
  Dim FiltroTipo As String, FiltroBanco As String, FiltroPeriodo As String
  
  FrmRelChEmitidos.MousePointer = vbHourglass
  XGT_CONJUNCAO = ""
  XGT_SELECAO = ""
  
  'Período
  If (Not IsNull(dtpInicio.Value)) And (Not IsNull(dtpFim.Value)) Then
    XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsFINChequesEmitidos.dupl_dt_pagrec >= '" & Format(dtpInicio.Value, "mm/dd/yyyy") & _
      "' AND ConsFINChequesEmitidos.dupl_dt_pagrec <= '" & Format(dtpFim.Value, "mm/dd/yyyy") & "'"
    FiltroPeriodo = "Período: " & Format$(dtpInicio.Value, "dd/mm/yy") & " a " & Format$(dtpFim.Value, "dd/mm/yy")
    XGT_CONJUNCAO = " AND "
  ElseIf (Not IsNull(dtpInicio.Value)) And (IsNull(dtpFim.Value)) Then
    XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsFINChequesEmitidos.dupl_dt_pagrec >= '" & Format(dtpInicio.Value, "mm/dd/yyyy") & "'"
    FiltroPeriodo = "Período: A partir de " & Format$(dtpInicio.Value, "dd/mm/yy")
    XGT_CONJUNCAO = " AND "
  ElseIf (IsNull(dtpInicio.Value)) And (Not IsNull(dtpFim.Value)) Then
    XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsFINChequesEmitidos.dupl_dt_pagrec <= '" & Format(dtpFim.Value, "mm/dd/yyyy") & "'"
    FiltroPeriodo = "Período: Até " & Format$(dtpFim.Value, "dd/mm/yy")
    XGT_CONJUNCAO = " AND "
  Else
    FiltroPeriodo = "Período: Geral"
  End If
  
  'Banco
  If Trim$(CboCCorrente2.Text) <> "" Then
    DatBanco.Recordset.bookmark = CboCCorrente2.SelectedItem
    XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsFINChequesEmitidos.coco_cd_codigo = " & CStr(CboCCorrente2.BoundText)
    FiltroBanco = "Banco: " & Trim$(DatBanco.Recordset.Fields("banc_tx_descricao")) & " Agência: " & _
      Trim$(DatBanco.Recordset.Fields("coco_nm_agencia")) & " Conta: " & _
      Trim$(DatBanco.Recordset.Fields("coco_tx_conta")) & ""
    XGT_CONJUNCAO = " AND "
  Else
    FiltroBanco = "Banco, Agência e Conta: Geral"
  End If
  
  'Filtra a empresa Ativa
  XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "ConsFINChequesEmitidos.empr_cd_empresa = " & CStr(PCodEmpresa) + ""

  'Abre os recordsets e verificar se eles possuem dados
  Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINChequesEmitidos WHERE " & XGT_SELECAO)
  If XGO_RSRELATORIO.EOF Then
    MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
    FrmRelChEmitidos.MousePointer = vbDefault
    Exit Sub
  End If
  Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas")
                  
  'Carrega o Relatório
  FrmVisRelatorios.Caption = "Relação de Cheques Emitidos"
  Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.path + "\Relatorios\RelChequesEmitidos.rpt")
  
  'Passa o recordset para o relatório
  XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO, 3, 1
  
  'Ativa os Relatórios
  With XGR_RELATORIO.FormulaFields
    .GetItemByName("Periodo").Text = Chr(34) + FiltroPeriodo + Chr(34)
    .GetItemByName("Banco").Text = Chr(34) + FiltroBanco + Chr(34)
    .GetItemByName("CodEmpresa").Text = Chr(34) + CStr(PCodEmpresa) + Chr(34)
  End With
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
  FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
  FrmRelChEmitidos.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
  Set Formulario = FrmRelChEmitidos
  subManutencaoJanelasAtivas "I", "FrmRelChEmitidos"
  'Alteração filtro CC por CGC Raiz - Samuel 27/09/18
  If Mid(PEmpresa, 1, 10) = "QUALIDADOS" Or Mid(PEmpresa, 1, 8) = "QUALINFO" Then
    subConectarControleDadosNV DatBanco, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_CGCRaiz = '" & Mid(XGT_CGC, 1, 10) & "' ORDER BY coco_tx_Descricao", Estatico
  Else
    subConectarControleDadosNV DatBanco, "SELECT * FROM ConsGENCCcombo where empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
  End If
End Sub
