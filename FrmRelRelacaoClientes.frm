VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmRelRelacaoClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório Relação de Clientes"
   ClientHeight    =   4410
   ClientLeft      =   3330
   ClientTop       =   3615
   ClientWidth     =   6630
   Icon            =   "FrmRelRelacaoClientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraOrdem 
      Caption         =   "Ordem"
      Enabled         =   0   'False
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
      Left            =   900
      TabIndex        =   18
      Top             =   3060
      Width           =   4695
      Begin VB.OptionButton OptCliente 
         Caption         =   "Cliente"
         Enabled         =   0   'False
         Height          =   195
         Left            =   780
         TabIndex        =   20
         Top             =   300
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton OptImovel 
         Caption         =   "Imóvel"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3000
         TabIndex        =   19
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   900
      TabIndex        =   14
      Top             =   1680
      Width           =   4695
      Begin VB.CheckBox ChkDadosContrato 
         Caption         =   "Apresentar dados do Contrato?"
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
         Left            =   360
         TabIndex        =   17
         Top             =   180
         Value           =   1  'Checked
         Width           =   3075
      End
      Begin VB.CheckBox ChkEmpreendimentos 
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
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   960
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.CheckBox ChkDadosCliente 
         Caption         =   "Apresentar dados do Cliente?"
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
         Left            =   360
         TabIndex        =   15
         Top             =   600
         Value           =   1  'Checked
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   6495
      Begin MSDataListLib.DataCombo CboCliente 
         Bindings        =   "FrmRelRelacaoClientes.frx":2AFA
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   540
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "focl_tx_RazaoSocial"
         BoundColumn     =   "focl_cd_FornCli"
         Text            =   "CboCliente"
      End
      Begin Threed.SSCommand CmdLimpar 
         Height          =   315
         Left            =   6060
         TabIndex        =   8
         Top             =   540
         Width           =   315
         _Version        =   65536
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   78
         MouseIcon       =   "FrmRelRelacaoClientes.frx":2B14
         Picture         =   "FrmRelRelacaoClientes.frx":2B30
      End
      Begin MSMask.MaskEdBox MskEmpreendimento 
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   120
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
         Bindings        =   "FrmRelRelacaoClientes.frx":2C42
         Height          =   315
         Left            =   2160
         TabIndex        =   10
         Top             =   120
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
         Left            =   6060
         TabIndex        =   11
         Top             =   120
         Width           =   315
         _Version        =   65536
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   78
         MouseIcon       =   "FrmRelRelacaoClientes.frx":2C62
         Picture         =   "FrmRelRelacaoClientes.frx":2C7E
      End
      Begin VB.Label LblCliente 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente:"
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
         Left            =   720
         TabIndex        =   13
         Top             =   600
         Width           =   765
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
         Left            =   60
         TabIndex        =   12
         Top             =   180
         Width           =   1485
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   5220
      TabIndex        =   4
      Top             =   3960
      Width           =   1035
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir..."
      Height          =   345
      Left            =   4110
      TabIndex        =   3
      Top             =   3960
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc DatClientes 
      Height          =   330
      Left            =   1740
      Top             =   3900
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
      Caption         =   "DatCliente"
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
   Begin MSAdodcLib.Adodc DatEmpreendimento 
      Height          =   330
      Left            =   480
      Top             =   3900
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
   Begin Threed.SSFrame FraStatus 
      Height          =   660
      Left            =   900
      TabIndex        =   5
      Top             =   1020
      Width           =   4725
      _Version        =   65536
      _ExtentX        =   8334
      _ExtentY        =   1164
      _StockProps     =   14
      Caption         =   "Status"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         Left            =   1860
         TabIndex        =   1
         Top             =   300
         Value           =   1  'Checked
         Width           =   1080
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
         Left            =   450
         TabIndex        =   0
         Top             =   300
         Value           =   1  'Checked
         Width           =   1065
      End
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
         Left            =   3330
         TabIndex        =   2
         Top             =   300
         Value           =   1  'Checked
         Width           =   1230
      End
   End
End
Attribute VB_Name = "FrmRelRelacaoClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XFO_RSCLIENTESIMOVEIS As New ADODB.Recordset

Private Sub CboCliente_Click(Area As Integer)
    If CboCliente.Text <> "" Then
        DatClientes.Recordset.bookmark = CboCliente.SelectedItem
    End If
End Sub

Private Sub CboEmpreendimento_Change()
    If CboEmpreendimento.Text <> "" Then
      DatEmpreendimento.Recordset.bookmark = CboEmpreendimento.SelectedItem
      MskEmpreendimento.Text = CboEmpreendimento.BoundText
      subConectarControleDadosNV DatClientes, "SELECT focl_cd_FornCli,focl_tx_RazaoSocial FROM ConsCAPImoveisPorCliente " & _
        " WHERE empr_cd_Empresa= " & PCodEmpresa & " AND empd_cd_Empreendimento='" & CboEmpreendimento.BoundText & "'" & _
        " ORDER BY focl_tx_RazaoSocial", Estatico
    Else
      subConectarControleDadosNV DatClientes, "SELECT focl_cd_FornCli,focl_tx_RazaoSocial " & _
      "FROM ConsCAPImoveisPorCliente WHERE empr_cd_Empresa= " & PCodEmpresa & " ORDER BY focl_tx_RazaoSocial", Estatico
    End If
    CboCliente.BoundText = ""
    
End Sub

Private Sub ChkEmpreendimentos_Click()
  
  If ChkEmpreendimentos.Value = 1 Then
    FraOrdem.Enabled = False
    OptCliente.Enabled = False
    OptImovel.Enabled = False
  Else
    FraOrdem.Enabled = True
    OptCliente.Enabled = True
    OptImovel.Enabled = True
  End If
  
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
    Dim XLT_FILTROSTATUS As String  'Armazenar filtro do status p/ relatório
    Dim XLT_FILTROPERIODO As String 'Armazenar filtro do período escolhido
    Dim XLT_FILTRODADOSCONTRATO As String 'Armazenar filtro Dados do contrato
    Dim XLT_STATUS As String        '1º Status escolhido
    Dim XLT_STATUS2 As String        '2º Status escolhido
    Dim XLT_EMPREENDIMENTO As String 'Empreendimento escolhido
    Dim XLT_CLIENTE As String 'Cliente escolhido
    Dim XLT_FILTRODADOSCLIENTE As String 'Armazenar filtro Dados do contrato
    
    FrmRelRelacaoClientes.MousePointer = vbHourglass
 
    XGT_SELECAO = ""
    XGT_CONJUNCAO = ""
                  
    'Empreendimento
    If CboEmpreendimento.BoundText <> "" Then
      XLT_EMPREENDIMENTO = CboEmpreendimento.BoundText
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsCAPRelRelacaoClientes.empd_cd_Empreendimento = '" & XLT_EMPREENDIMENTO & "'"
      XGT_CONJUNCAO = " AND "
    End If
    
   'Cliente
    If CboCliente.BoundText <> "" Then
      XLT_CLIENTE = CboCliente.BoundText
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsCAPRelRelacaoClientes.focl_cd_FornCli = '" & XLT_CLIENTE & "'"
      XGT_CONJUNCAO = " AND "
    End If
  
    'Status
    If ChkAtivo.Value = ChkInativo.Value And ChkAtivo.Value = ChkCedido.Value Then
        XLT_FILTROSTATUS = "Status: Geral"
        XLT_STATUS = "T" 'Todos
    Else
        If ChkAtivo.Value = 1 Then
            If ChkInativo.Value = 1 Then  'Ativo e Inativo
                XLT_STATUS = "ATIVO"
                XLT_STATUS2 = "INATIVO"
                XLT_FILTROSTATUS = "Status: Ativo e Inativo"
            ElseIf ChkCedido.Value = 1 Then  'Ativo e Cedido
                XLT_STATUS = "ATIVO"
                XLT_STATUS2 = "CEDIDO"
                XLT_FILTROSTATUS = "Status: Ativo e Cedido"
            Else
                XLT_STATUS = "ATIVO" 'Ativo
                XLT_FILTROSTATUS = "Status: Ativo"
            End If
        ElseIf ChkInativo.Value = 1 Then
            If ChkCedido.Value = 1 Then   'Inativo e Cedido
                XLT_STATUS = "INATIVO"
                XLT_STATUS2 = "CEDIDO"
                XLT_FILTROSTATUS = "Status: Inativo e Cedidod"
            Else
                XLT_STATUS = "INATIVO" 'Inativo
                XLT_FILTROSTATUS = "Status: Inativo"
            End If
        ElseIf ChkCedido.Value = 1 Then
            XLT_STATUS = "CEDIDO" 'Cedido
            XLT_FILTROSTATUS = "Status: Cedido"
        End If
        If XLT_STATUS2 <> "" Then 'Foi escolhido dois status
            XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "(ConsCAPRelRelacaoClientes.Status = '" & XLT_STATUS & "'" & _
            " OR ConsCAPRelRelacaoClientes.Status = '" & XLT_STATUS2 & "')"
            XGT_CONJUNCAO = " AND "
        Else
            XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "ConsCAPRelRelacaoClientes.Status = '" & XLT_STATUS & "'"
            XGT_CONJUNCAO = " AND "
        End If
    End If
            
    'Filtra a empresa Ativa
    XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + "ConsCAPRelRelacaoClientes.empr_cd_empresa = " & CStr(PCodEmpresa)
    
    XGT_SELECAO = XGT_SELECAO + " ORDER BY focl_tx_RazaoSocial"
    
    'Filtro Dados Contrato
    If ChkDadosContrato.Value = 1 Then
      XLT_FILTRODADOSCONTRATO = "S"
    Else
      XLT_FILTRODADOSCONTRATO = "N"
    End If
    
    'Filtro Dados Cliente
    If ChkDadosCliente.Value = 1 Then
      XLT_FILTRODADOSCLIENTE = "S"
    Else
      XLT_FILTRODADOSCLIENTE = "N"
    End If
    
    'Fecha o recordset se estiver aberto
    If XGO_RSRELATORIO.State = adStateOpen Then
      XGO_RSRELATORIO.Close
      Set XGO_RSRELATORIO = Nothing
    End If
    'Abre os recordsets e verifica se eles possuem dados
    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsCAPRelRelacaoClientes WHERE " & XGT_SELECAO)
    
    If XGO_RSRELATORIO.EOF Then
        MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
        FrmRelRelacaoClientes.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Fecha o recordset se estiver aberto
    If XGO_RSCABECALHO.State = adStateOpen Then
      XGO_RSCABECALHO.Close
      Set XGO_RSCABECALHO = Nothing
    End If
    Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
  
    'Carrega os relatórios
    FrmVisRelatorios.Caption = "Relatório de Relação dos Clientes"
    If ChkEmpreendimentos.Value = 1 Then
      Set XGR_RELATORIO = Nothing
      Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelRelacaoClientesAgrupados.rpt")
    Else
      Set XGR_RELATORIO = Nothing
      If OptCliente.Value = True Then
        Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelRelacaoClientes.rpt")
      Else
        Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelRelacaoClientesPorImovel.rpt")
      End If
      
      'Fecha o recordset se estiver aberto
      If XFO_RSCLIENTESIMOVEIS.State = adStateOpen Then
        XFO_RSCLIENTESIMOVEIS.Close
        Set XFO_RSCLIENTESIMOVEIS = Nothing
      End If
      Set XFO_RSCLIENTESIMOVEIS = ConexaoRelatorio.Execute("SELECT * FROM ConsCapRelClientesImoveis WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY imov_cd_Imovel")
    End If
  
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
  
    'Ativa os Relatórios
    With XGR_RELATORIO.FormulaFields
        '.GetItemByName("Periodo").Text = Chr(34) & XLT_FILTROPERIODO & Chr(34)
        .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
        .GetItemByName("Status").Text = Chr(34) & XLT_FILTROSTATUS & Chr(34)
        .GetItemByName("DadosContrato").Text = Chr(34) & XLT_FILTRODADOSCONTRATO & Chr(34)
        .GetItemByName("DadosCliente").Text = Chr(34) & XLT_FILTRODADOSCLIENTE & Chr(34)
        
    End With
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
    If ChkEmpreendimentos.Value = 0 Then
      FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XFO_RSCLIENTESIMOVEIS, "ClientesImoveis"
    End If
    FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
    FrmRelRelacaoClientes.MousePointer = vbDefault
End Sub

Private Sub CmdLimpar_Click()
    CboCliente.Text = ""
End Sub

Private Sub CmdLimparEmpreendimento_Click()
    CboEmpreendimento.BoundText = ""
    MskEmpreendimento.Text = ""
End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelRelacaoClientes
End Sub

Private Sub Form_Load()
    subManutencaoJanelasAtivas "I", "FrmRelRelacaoClientes"
    'subConectarControleDadosNV DatClientes, "SELECT focl_cd_forncli,focl_tx_classe,focl_tx_razaoSocial FROM FornClientes WHERE focl_tx_classe = 'C' ORDER BY focl_tx_razaoSocial", Estatico
    subConectarControleDadosNV DatClientes, "SELECT focl_cd_FornCli,focl_tx_RazaoSocial " & _
      "FROM ConsCAPImoveisPorCliente WHERE empr_cd_Empresa= " & PCodEmpresa & " ORDER BY focl_tx_RazaoSocial", Estatico
    subConectarControleDadosNV DatEmpreendimento, "SELECT * FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_tx_Nome", Estatico
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  If XGO_RSRELATORIO.State = adStateOpen Then
    XGO_RSRELATORIO.Close
  End If
  
  If XGO_RSCABECALHO.State = adStateOpen Then
    XGO_RSCABECALHO.Close
  End If
  
  If XFO_RSCLIENTESIMOVEIS.State = adStateOpen Then
    XFO_RSCLIENTESIMOVEIS.Close
  End If
  
  Set XGO_RSRELATORIO = Nothing
  Set XGO_RSCABECALHO = Nothing
  Set XFO_RSCLIENTESIMOVEIS = Nothing
  
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
