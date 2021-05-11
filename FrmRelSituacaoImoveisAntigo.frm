VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRelSituacaoImoveis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório Situação dos Imóveis"
   ClientHeight    =   4260
   ClientLeft      =   3330
   ClientTop       =   3615
   ClientWidth     =   6435
   Icon            =   "FrmRelSituacaoImoveis3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraGaragens 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   840
      TabIndex        =   15
      Top             =   3000
      Width           =   4710
      Begin VB.CheckBox ChkGaragens 
         Caption         =   "Apresentar garagens?"
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
         TabIndex        =   16
         Top             =   180
         Width           =   3195
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
      Left            =   840
      TabIndex        =   6
      Top             =   720
      Width           =   4710
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   795
         TabIndex        =   7
         Top             =   285
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yy"
         DateIsNull      =   -1  'True
         Format          =   3801091
         CurrentDate     =   37687.5817592593
      End
      Begin MSComCtl2.DTPicker dtpFim 
         Height          =   315
         Left            =   3000
         TabIndex        =   8
         Top             =   285
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yy"
         DateIsNull      =   -1  'True
         Format          =   3801091
         CurrentDate     =   37687.5814814815
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   330
         Width           =   495
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   4530
      TabIndex        =   1
      Top             =   3795
      Width           =   1035
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir..."
      Height          =   345
      Left            =   3390
      TabIndex        =   0
      Top             =   3795
      Width           =   1035
   End
   Begin MSMask.MaskEdBox MskEmpreendimento 
      Height          =   315
      Left            =   1575
      TabIndex        =   2
      Top             =   270
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
      Bindings        =   "FrmRelSituacaoImoveis3.frx":2AFA
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   270
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
   Begin MSAdodcLib.Adodc DatEmpreendimento 
      Height          =   330
      Left            =   900
      Top             =   3780
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
   Begin Threed.SSCommand CmdLimpar 
      Height          =   315
      Left            =   6030
      TabIndex        =   5
      Top             =   255
      Width           =   315
      _Version        =   65536
      _ExtentX        =   556
      _ExtentY        =   556
      _StockProps     =   78
      MouseIcon       =   "FrmRelSituacaoImoveis3.frx":2B1A
      Picture         =   "FrmRelSituacaoImoveis3.frx":2B36
   End
   Begin Threed.SSFrame FraStatus 
      Height          =   660
      Left            =   825
      TabIndex        =   11
      Top             =   1530
      Width           =   4725
      _Version        =   65536
      _ExtentX        =   8334
      _ExtentY        =   1164
      _StockProps     =   14
      Caption         =   "Status do Imóvel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CheckBox ChkVendido 
         Caption         =   "Vendido"
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
         Left            =   1935
         TabIndex        =   13
         Top             =   300
         Value           =   1  'Checked
         Width           =   1080
      End
      Begin VB.CheckBox ChkDisponivel 
         Caption         =   "Disponível"
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
         Left            =   450
         TabIndex        =   12
         Top             =   300
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox ChkReservado 
         Caption         =   "Reservado"
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
         Left            =   3330
         TabIndex        =   14
         Top             =   300
         Value           =   1  'Checked
         Width           =   1230
      End
   End
   Begin Threed.SSFrame FraStatusContrato 
      Height          =   660
      Left            =   840
      TabIndex        =   17
      Top             =   2280
      Width           =   4725
      _Version        =   65536
      _ExtentX        =   8334
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
         Left            =   3000
         TabIndex        =   20
         Top             =   300
         Width           =   1230
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
         TabIndex        =   19
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
         Left            =   1740
         TabIndex        =   18
         Top             =   300
         Width           =   1080
      End
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
      Left            =   75
      TabIndex        =   4
      Top             =   315
      Width           =   1485
   End
End
Attribute VB_Name = "FrmRelSituacaoImoveis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CboEmpreendimento_Change()
    If CboEmpreendimento.Text <> "" Then
        DatEmpreendimento.Recordset.bookmark = CboEmpreendimento.SelectedItem
        MskEmpreendimento.Text = CboEmpreendimento.BoundText
    End If
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
    Dim XLT_FILTROSTATUS As String  'Armazenar filtro do status p/ relatório
    Dim XLT_FILTROSTATUSCONTRATO As String  'Armazenar filtro do status p/ relatório
    Dim XLT_FILTROPERIODO As String 'Armazenar filtro do período escolhido
    Dim XLT_STATUS As String        '1º Status escolhido
    Dim XLT_STATUS2 As String        '2º Status escolhido
    Dim XLT_EMPREENDIMENTO As String 'Empreendimento escolhido
    
    FrmRelSituacaoImoveis.MousePointer = vbHourglass
 
    XGT_SELECAO = ""
    XGT_CONJUNCAO = ""
  
    'Período
    If (Not IsNull(dtpInicio.Value)) And (Not IsNull(dtpFim.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelSituacaoImoveis.dtVenda >= " & FunNuloData(dtpInicio.Value, NomeSgbd) & "" & _
        " AND ConsCAPRelSituacaoImoveis.dtVenda <= " & FunNuloData(dtpFim.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODO = "Período da Venda: " & Format$(dtpInicio.Value, "dd/mm/yy") & " a " & Format$(dtpFim.Value, "dd/mm/yy")
    ElseIf (Not IsNull(dtpInicio.Value)) And (IsNull(dtpFim.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelSituacaoImoveis.dtVenda >= " & FunNuloData(dtpInicio.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODO = "Período da Venda: A partir de " & Format$(dtpInicio.Value, "dd/mm/yy")
    ElseIf (IsNull(dtpInicio.Value)) And (Not IsNull(dtpFim.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelSituacaoImoveis.dtVenda <= " & FunNuloData(dtpFim.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODO = "Período da Venda: Até " & Format$(dtpFim.Value, "dd/mm/yy")
    Else
        XLT_FILTROPERIODO = "Período da Venda: Geral"
    End If
    
    'Empreendimento
      If CboEmpreendimento.BoundText <> "" Then
        XLT_EMPREENDIMENTO = CboEmpreendimento.BoundText
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelSituacaoImoveis.empd_cd_Empreendimento = '" & XLT_EMPREENDIMENTO & "'"
        XGT_CONJUNCAO = " AND "
      End If
                  
    'Status do Imóvel
    If ChkDisponivel.Value = ChkVendido.Value And ChkDisponivel.Value = ChkReservado.Value Then
        XLT_FILTROSTATUS = "Status: Geral"
        XLT_STATUS = "T" 'Todos
    Else
        If ChkDisponivel.Value = 1 Then
            If ChkVendido.Value = 1 Then  'Disponível e Vendido
                XLT_STATUS = "D"
                XLT_STATUS2 = "V"
                XLT_FILTROSTATUS = "Status: Disponível e Vendido"
            ElseIf ChkReservado.Value = 1 Then  'Disponível e Reservado
                XLT_STATUS = "D"
                XLT_STATUS2 = "R"
                XLT_FILTROSTATUS = "Status: Disponível e Reservado"
            Else
                XLT_STATUS = "D" 'Disponível
                XLT_FILTROSTATUS = "Status: Disponível"
            End If
        ElseIf ChkVendido.Value = 1 Then
            If ChkReservado.Value = 1 Then   'Vendido e Reservado
                XLT_STATUS = "V"
                XLT_STATUS2 = "R"
                XLT_FILTROSTATUS = "Status: Vendido e Reservado"
            Else
                XLT_STATUS = "V" 'Vendido
                XLT_FILTROSTATUS = "Status: Vendido"
            End If
        ElseIf ChkReservado.Value = 1 Then
            XLT_STATUS = "R" 'Reservado
            XLT_FILTROSTATUS = "Status: Reservado"
        End If
        If XLT_STATUS2 <> "" Then 'Foi escolhido dois status
            XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "(imov_tx_status = '" & XLT_STATUS & "'" & _
            " OR imov_tx_Status = '" & XLT_STATUS2 & "')"
        Else
            XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "imov_tx_status = '" & XLT_STATUS & "'"
        End If
        XGT_CONJUNCAO = " AND "
    End If
    
    XLT_STATUS = ""
    XLT_STATUS2 = ""
    'Status do Contrato
    If ChkAtivo.Value = ChkInativo.Value And ChkAtivo.Value = ChkCedido.Value Then
        XLT_FILTROSTATUS = "Status: Geral"
        XLT_STATUS = "T" 'Todos
    Else
        If ChkAtivo.Value = 1 Then
            If ChkInativo.Value = 1 Then  'Ativo e Inativo
                XLT_STATUS = "A"
                XLT_STATUS2 = "I"
                XLT_FILTROSTATUSCONTRATO = "Status: Ativo e Inativo"
            ElseIf ChkCedido.Value = 1 Then  'Ativo e Cedido
                XLT_STATUS = "A"
                XLT_STATUS2 = "C"
                XLT_FILTROSTATUSCONTRATO = "Status: Ativo e Cedido"
            Else
                XLT_STATUS = "A" 'Ativo
                XLT_FILTROSTATUSCONTRATO = "Status: Ativo"
            End If
        ElseIf ChkInativo.Value = 1 Then
            If ChkCedido.Value = 1 Then   'Inativo e Cedido
                XLT_STATUS = "I"
                XLT_STATUS2 = "C"
                XLT_FILTROSTATUSCONTRATO = "Status: Inativo e Cedidod"
            Else
                XLT_STATUS = "I" 'Inativo
                XLT_FILTROSTATUSCONTRATO = "Status: Inativo"
            End If
        ElseIf ChkCedido.Value = 1 Then
            XLT_STATUS = "C" 'Cedido
            XLT_FILTROSTATUSCONTRATO = "Status: Cedido"
        End If
        If XLT_STATUS2 <> "" Then 'Foi escolhido dois status
            XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "(cont_tx_Status = '" & XLT_STATUS & "'" & _
            " OR cont_tx_status = '" & XLT_STATUS2 & "')"
            XGT_CONJUNCAO = " AND "
        Else
            XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "cont_tx_Status = '" & XLT_STATUS & "'"
            XGT_CONJUNCAO = " AND "
        End If
    End If
            
    'Filtra a empresa Ativa
    XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + "ConsCAPRelSituacaoImoveis.empr_cd_empresa = " & CStr(PCodEmpresa)
  
    'Ordem
    XGT_SELECAO = XGT_SELECAO & " ORDER BY " & _
      "ConsCAPRelSituacaoImoveis.empd_cd_empreendimento, ConsCAPRelSituacaoImoveis.imov_cd_Imovel"
    
  
    'Abre os recordsets e verifica se eles possuem dados
    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsCAPRelSituacaoImoveis WHERE " & XGT_SELECAO)
    If XGO_RSRELATORIO.EOF Then
        MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
        FrmRelSituacaoImoveis.MousePointer = vbDefault
        Exit Sub
    End If
    Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
  
    'Carrega os relatórios
    FrmVisRelatorios.Caption = "Relatório de Situação dos Imóveis"
    
    If ChkGaragens.Value = 1 Then
      Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelSituacaoImovelGaragens.rpt")
    Else
      Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelSituacaoImovel.rpt")
    End If
  
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
                            
    'Ativa os Relatórios
    With XGR_RELATORIO.FormulaFields
        .GetItemByName("Periodo").Text = Chr(34) & XLT_FILTROPERIODO & Chr(34)
        .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
        .GetItemByName("Status").Text = Chr(34) & XLT_FILTROSTATUS & Chr(34)
    End With
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
    FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
    FrmRelSituacaoImoveis.MousePointer = vbDefault
End Sub

Private Sub CmdLimpar_Click()
    CboEmpreendimento.Text = ""
    MskEmpreendimento.Text = ""
End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelSituacaoImoveis
End Sub

Private Sub Form_Load()
  subManutencaoJanelasAtivas "I", "FrmRelSituacaoImoveis"
  subConectarControleDadosNV DatEmpreendimento, "SELECT * FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_tx_Nome", Estatico
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


